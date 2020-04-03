"use strict";

const execSync = require('child_process').execSync;
const puppeteer = require("puppeteer");
const term = require("terminal-kit").terminal;
const fs = require("fs");
const url = require('url');
const path = require("path");
const yargs = require("yargs");
var m3u8Parser = require("m3u8-parser");
const request = require('request');

const argv = yargs.options({
    v: { alias:'videoUrls', type: 'array', demandOption: true },
    u: { alias:'username', type: 'string', demandOption: false, describe: 'Your Microsoft Email' },
    p: { alias:'password', type: 'string', demandOption: false },
    o: { alias:'outputDirectory', type: 'string', default: 'videos' },
    q: { alias: 'quality', type: 'number', demandOption: false, describe: 'Video Quality, usually [0-5]'},
    p: { alias: 'polimi', type: 'boolean', default: false, demandOption: false, describe: 'Use PoliMi Login. If set, use Codice Persona as username'},
    k: { alias: 'noKeyring', type: 'boolean', default: false, demandOption: false, describe: 'Do not use system keyring'},
    c: { alias: 'conn', type: 'number', default: 16, demandOption: false, describe: 'Number of simultaneous connections [1-16]'}
})
.help('h')
.alias('h', 'help')
.example('node $0 -v "https://web.microsoftstream.com/video/9611baf5-b12e-4782-82fb-b2gf68c05adc"\n', "Standard usage")
.example('node $0 -p -v "https://web.microsoftstream.com/video/9611baf5-b12e-4782-82fb-b2gf68c05adc"\n', "Use PoliMiLogin")
.example('node $0 -v "https://web.microsoftstream.com/video/9611baf5-b12e-4782-82fb-b2gf68c05adc" "https://web.microsoftstream.com/video/6711baa5-c56e-4782-82fb-c2ho68c05zde"\n', "Multiple videos download")
.example('node $0 -v "https://web.microsoftstream.com/video/9611baf5-b12e-4782-82fb-b2gf68c05adc" -q 4\n', "Define default quality download to avoid manual prompt")
.example('node $0 -v "https://web.microsoftstream.com/video/9611baf5-b12e-4782-82fb-b2gf68c05adc" -o "C:\\Lessons\\Videos"\n', "Define output directory (absoulte o relative path)")
.example('node $0 -u EMAIL -p PASSWORD -v "https://web.microsoftstream.com/video/9611baf5-b12e-4782-82fb-b2gf68c05adc"\n', "Replace saved email and/or password")
.example('node $0 -v "https://web.microsoftstream.com/video/9611baf5-b12e-4782-82fb-b2gf68c05adc" -k\n', "Do not save the password into system keyring")
.example('node $0 -c 10 -v "https://web.microsoftstream.com/video/9611baf5-b12e-4782-82fb-b2gf68c05adc" -k\n', "Define number of simultaneous connections. Reduce it if you encounter problems during downloading")
.argv;

console.info('\nVideo URLs: %s', argv.videoUrls);
if(typeof argv.username !== 'undefined') {console.info('Email: %s', argv.username);}
console.info('Output Directory: %s\n', argv.outputDirectory);

function sanityChecks() {
    try {
        const aria2Ver = execSync('aria2c --version').toString().split('\n')[0];
        term.green(`Using ${aria2Ver}\n`);
    }
    catch (e) {
        term.red('You need aria2c in $PATH for this to work. Make sure it is a relatively recent one.');
        process.exit(22);
    }
    try {
        const ffmpegVer = execSync('ffmpeg -version').toString().split('\n')[0];
        term.green(`Using ${ffmpegVer}\n\n`);
    }
    catch (e) {
        term.red('FFmpeg is missing. You need a fairly recent release of FFmpeg in $PATH.');
        process.exit(23);
    }
    if (!fs.existsSync(argv.outputDirectory)) {
        if (path.isAbsolute(argv.outputDirectory) || argv.outputDirectory[0] == '~') console.log('Creating output directory: ' + argv.outputDirectory);
        else console.log('Creating output directory: ' + process.cwd() + path.sep + argv.outputDirectory);
        try {
          fs.mkdirSync(argv.outputDirectory, { recursive: true }); // use native API for nested directory. No recursive function needed, but compatible only with node v10 or later
        } catch (e) {
          term.red("Can not create nested directories. Node v10 or later is required\n");
          process.exit();
        }
    }

}


async function downloadVideo(videoUrls, email, password, outputDirectory) {

   email = await handleEmail(email);

   // handle password
   const keytar = require('keytar');
   //keytar.deletePassword('MStreamDownloader', email);
   if(password === null) { // password not passed as argument
        var password = {};
        if(argv.noKeyring === false) {
          try {
              await keytar.getPassword("MStreamDownloader", email).then(function(result) { password = result; });
              if (password === null) { // no previous password saved
                  password = await promptPassword("Password not saved. Please enter your password, MStreamDownloader will not ask for it next time: ");
                  await keytar.setPassword("MStreamDownloader", email, password);
              } else {
                  console.log("\nReusing password saved in system's keychain!")
              }
          }
          catch(e) {
              console.log("X11 is not installed on this system. MStreamDownloader can't use keytar to save the password.")
              password = await promptPassword("No problem, please manually enter your password: ");
          }
        } else {
          password = await promptPassword("\nPlease enter your password: ");
        }
   } else {
      if(argv.noKeyring === false) {
        try {
            await keytar.setPassword("MStreamDownloader", email, password);
            console.log("Your password has been saved. Next time, you can avoid entering it!");
        } catch(e) {
            // X11 is missing. Can't use keytar
        }
      }
   }
   console.log('\nLaunching headless Chrome to perform the OpenID Connect dance...');
   const browser = await puppeteer.launch({
       // Switch to false if you need to login interactively
       headless: true,
       args: ['--disable-dev-shm-usage', '--lang=it-IT']
   });

   const page = await browser.newPage();
   console.log('Navigating to STS login page...');
   await page.goto('https://web.microsoftstream.com/', { waitUntil: 'networkidle2' });

   if(argv.polimi === true) {
       await polimiLogin(page, email, password)
   } else {
       await defaultLogin(page, email, password)
   }

   await browser.waitForTarget(target => target.url().includes('microsoftstream.com/'), { timeout: 90000 });
   console.log('We are logged in. ');
   await sleep (3000)
   const cookie = await extractCookies(page);
   console.log('Got required authentication cookies.');
    for (let videoUrl of videoUrls) {
       term.green(`\nStart downloading video: ${videoUrl}\n`);

       var videoID = videoUrl.substring(videoUrl.indexOf("/video/")+7, videoUrl.length).substring(0, 36); // use the video id (36 character after '/video/') as temp dir name
       var full_tmp_dir = path.join(argv.outputDirectory, videoID);

       var headers = {
           'Cookie': cookie
       };

       var options = {
           url: 'https://euwe-1.api.microsoftstream.com/api/videos/'+videoID+'?api-version=1.0-private',
           headers: headers
       };
       var response = await doRequest(options);
       const obj = JSON.parse(response);

       if(obj.hasOwnProperty('error')) {
         let errorMsg = ''
         if(obj.error.code === 'Forbidden') {
           errorMsg = 'You are not authorized to access this video.\n'
         } else {
           errorMsg = '\nError downloading this video.\n'
         }
         term.red(errorMsg)
         continue;
       }

       // creates tmp dir
       if (!fs.existsSync(full_tmp_dir)) {
           fs.mkdirSync(full_tmp_dir);
       } else {
           rmDir(full_tmp_dir);
           fs.mkdirSync(full_tmp_dir);
       }

       var title = (obj.name).trim();
       console.log(`\nVideo title is: ${title}`);
       title = title.replace(/[/\\?%*:;|"<>]/g, '-'); // remove illegal characters
       var isoDate = obj.publishedDate;
       if (isoDate !== null && isoDate !== '') {
          let date = new Date(isoDate);
          let year = date.getFullYear();
          let month = date.getMonth()+1;
          let dt = date.getDate();

          if (dt < 10) {
            dt = '0' + dt;
          }
            if (month < 10) {
            month = '0' + month;
          }
          let uploadDate = dt + '_' + month + '_' + year;
          title = 'Lesson ' + uploadDate + ' - ' + title;
       } else {
            // console.log("no upload date found");
       }

      let playbackUrls = obj.playbackUrls
      let hlsUrl = ''
      for(var elem in playbackUrls) {
          if(playbackUrls[elem]['mimeType'] === 'application/vnd.apple.mpegurl') {
            var u = url.parse(playbackUrls[elem]['playbackUrl'], true);
            hlsUrl = u.query.playbackurl
            break;
          }
      }

        var options = {
            url: hlsUrl,
        };
        var response = await doRequest(options);
        var parser = new m3u8Parser.Parser();
        parser.push(response);
        parser.end();
        var parsedManifest = parser.manifest;

        var playlistsInfo = {};
        var question = '\n';
        var count = 0;
        var audioObj = null;
        var videoObj = null;
        for (var i=0 ; i<parsedManifest['playlists'].length ; i++) {
            if(parsedManifest['playlists'][i]['attributes'].hasOwnProperty('RESOLUTION')) {
                playlistsInfo[i] = {};
                playlistsInfo[i]['resolution'] =  parsedManifest['playlists'][i]['attributes']['RESOLUTION']['width'] + 'x' + parsedManifest['playlists'][i]['attributes']['RESOLUTION']['height'];
                playlistsInfo[i]['uri'] = parsedManifest['playlists'][i]['uri'];
                question = question + '[' + i + '] ' +  playlistsInfo[i]['resolution'] + '\n';
                count = count + 1;
            } else {
                 // if "RESOLUTION" key doesn't exist, means the current playlist is the audio playlist
                 // fix this for multiple audio tracks
                audioObj = parsedManifest['playlists'][i];
            }
        }
        //  if quality is passed as argument use that, otherwise prompt
        if (typeof argv.quality === 'undefined') {
            question = question + 'Choose the desired resolution: ';
            var res_choice = await promptResChoice(question, count);
        }
        else {
          if(argv.quality < 0 || argv.quality > count-1) {
            term.yellow(`Desired quality is not available for this video (available range: 0-${count-1})\nI'm going to use the best resolution available: ${playlistsInfo[count-1]['resolution']}`);
            var res_choice = count-1;
          }
          else {
            var res_choice = argv.quality;
            term.yellow(`Selected resolution: ${playlistsInfo[res_choice]['resolution']}`);
          }
        }

        videoObj = playlistsInfo[res_choice];

        const basePlaylistsUrl = hlsUrl.substring(0, hlsUrl.lastIndexOf("/") + 1);

        // **** VIDEO ****
        var videoLink = basePlaylistsUrl + videoObj['uri'];

        var headers = {
            'Cookie': cookie
        };
        var options = {
            url: videoLink,
            headers: headers
        };

        // *** Get protection key (same key for video and audio segments) ***
        var response = await doRequest(options);
        var parser = new m3u8Parser.Parser();
        parser.push(response);
        parser.end();
        var parsedManifest = parser.manifest;
        const keyUri = parsedManifest['segments'][0]['key']['uri'];
        var options = {
            url: keyUri,
            headers: headers,
            encoding: null
        };
        const key = await doRequest(options);

        var keyReplacement = '';
        if (path.isAbsolute(full_tmp_dir) || full_tmp_dir[0] == '~') { // absolute path
            var local_key_path = path.join(full_tmp_dir, 'my.key');
        }
        else {
            var local_key_path = path.join(process.cwd(), full_tmp_dir, 'my.key'); // requires absolute path in order to replace the URI inside the m3u8 file
        }
        fs.writeFileSync(local_key_path, key);
        if(process.platform === 'win32') {
          keyReplacement = await 'file:' + local_key_path.replace(/\\/g, '/');
        } else {
          keyReplacement = 'file://' + local_key_path;
        }


        // creates two m3u8 files:
        // - video_full.m3u8: to download all segements (replacing realtive segements path with absolute remote url)
        // - video_tmp.m3u8: used by ffmpeg to merge all downloaded segements (in this one we replace the remote key URI with the absoulte local path of the key downloaded above)
        var baseUri = videoLink.substring(0, videoLink.lastIndexOf("/") + 1);
        var video_full = await response.replace(new RegExp('Fragments', 'g'), baseUri+'Fragments'); // local path to full remote url path
        var video_tmp = await response.replace(keyUri, keyReplacement); // remote URI to local abasolute path
        var video_tmp = await video_tmp.replace(new RegExp('Fragments', 'g'), 'video_segments/Fragments');
        const video_full_path = path.join(full_tmp_dir, 'video_full.m3u8');
        const video_tmp_path = path.join(full_tmp_dir, 'video_tmp.m3u8');
        fs.writeFileSync(video_full_path, video_full);
        fs.writeFileSync(video_tmp_path, video_tmp);

        var n = argv.conn;
        if(n>16) n=16
        else if (n<1) n=1
        // download async. I'm Speed
        var aria2cCmd = 'aria2c -i "' + video_full_path + '" -j '+n+' -x '+n+' -d "' + path.join(full_tmp_dir, 'video_segments') + '" --header="Cookie:' + cookie + '"';
        var result = execSync(aria2cCmd, { stdio: 'inherit' });

        // **** AUDIO ****
        var audioLink = basePlaylistsUrl + audioObj['uri'];
        var options = {
            url: audioLink,
            headers: headers
        };

        // same as above but for audio segements
        var response = await doRequest(options);
        var baseUri = audioLink.substring(0, audioLink.lastIndexOf("/") + 1);
        var audio_full = await response.replace(new RegExp('Fragments', 'g'), baseUri+'Fragments');
        var audio_tmp = await response.replace(keyUri, keyReplacement);
        var audio_tmp = await audio_tmp.replace(new RegExp('Fragments', 'g'), 'audio_segments/Fragments');
        const audio_full_path = path.join(full_tmp_dir, 'audio_full.m3u8');
        const audio_tmp_path = path.join(full_tmp_dir, 'audio_tmp.m3u8');
        fs.writeFileSync(audio_full_path, audio_full);
        fs.writeFileSync(audio_tmp_path, audio_tmp);

        var aria2cCmd = 'aria2c -i "' + audio_full_path + '" -j '+n+' -x '+n+' -d "' + path.join(full_tmp_dir, 'audio_segments') + '" --header="Cookie:' + cookie + '"';
        var result = execSync(aria2cCmd, { stdio: 'inherit' });

        // *** MERGE audio and video segements in an mp4 file ***
        if (fs.existsSync(path.join(outputDirectory, title+'.mp4'))) {
            title = title + '-' + Date.now('nano');
        }

        // stupid Windows. Need to find a better way
        var ffmpegCmd = '';
        var ffmpegOpts = {stdio: 'inherit'};
        if(process.platform === 'win32') {
            ffmpegOpts['cwd'] = full_tmp_dir; // change working directory on windows, otherwise ffmpeg doesn't find the segements (relative paths problem, again, stupid windows. Or stupid me?)
            var outputFullPath = '';
            if (path.isAbsolute(outputDirectory) || outputDirectory[0] == '~')
              outputFullPath = path.join(outputDirectory, title);
            else
              outputFullPath = path.join('..', '..', outputDirectory, title);
            var ffmpegCmd = 'ffmpeg -protocol_whitelist file,http,https,tcp,tls,crypto,data -allowed_extensions ALL -i ' + 'audio_tmp.m3u8' + ' -protocol_whitelist file,http,https,tcp,tls,crypto,data -allowed_extensions ALL -i ' + 'video_tmp.m3u8' + ' -async 1 -c copy -bsf:a aac_adtstoasc -n "' + outputFullPath + '.mp4"';
        } else {
            var ffmpegCmd = 'ffmpeg -protocol_whitelist file,http,https,tcp,tls,crypto -allowed_extensions ALL -i "' + audio_tmp_path + '" -protocol_whitelist file,http,https,tcp,tls,crypto -allowed_extensions ALL -i "' + video_tmp_path + '" -async 1 -c copy -bsf:a aac_adtstoasc -n "' + path.join(outputDirectory, title) + '.mp4"';
        }

        var result = execSync(ffmpegCmd, ffmpegOpts);

        // remove tmp dir
        rmDir(full_tmp_dir);


    }

    console.log("\nAt this point Chrome's job is done, shutting it down...");
    await browser.close();
    term.green(`Done!\n`);

}

async function defaultLogin(page, email, password) {
    await page.waitForSelector('input[type="email"]');
    await page.keyboard.type(email);
    await page.click('input[type="submit"]');
    try {
      await page.waitForSelector('div[id="usernameError"]', { timeout: 1000 });
      term.red('Bad email');
      process.exit(401);
    } catch (error) {
       // email ok
    }

    // await sleep(2000) // maybe needed for slow connections
    await page.waitForSelector('input[type="password"]');
    await page.keyboard.type(password);
    await page.click('input[type="submit"]');

    try {
      await page.waitForSelector('div[id="passwordError"]', { timeout: 1000 });
      term.red('Bad password');
      process.exit(401);
    } catch (error) {
       // password ok
    }

    try {
      await page.waitForSelector('input[id="idBtn_Back"]', { timeout: 2000 });
      await page.click('input[id="idBtn_Back"]'); // Don't remember me
    } catch (error) {
       // button not appeared, ok...
    }
}

async function polimiLogin(page, username, password) {
    await page.waitForSelector('input[type="email"]');
    const tmpEmail = "11111111@polimi.it";
    await page.keyboard.type(tmpEmail);
    await page.click('input[type="submit"]');

    console.log('Filling in Servizi Online login form...');
    await page.waitForSelector('input[id="login"]');
    await page.type('input#login', username) // mette il codice persona
    await page.type('input#password', password) // mette la password
    await page.click('button[name="evn_conferma"]') // clicca sul tasto "Accedi"

    try {
      await page.waitForSelector('div[class="Message ErrorMessage"]', { timeout: 1000 });
      term.red('Bad credentials');
      process.exit(401);
    } catch (error) {
       // tutto ok
    }

    try {
        await page.waitForSelector('button[name="evn_continua"]', { timeout: 1000 }); // password is expiring
        await page.click('button[name="evn_continua"]');
    } catch (error) {
        // password is not expiring
    }

    try {
      await page.waitForSelector('input[id="idBtn_Back"]', { timeout: 2000 });
      await page.click('input[id="idBtn_Back"]'); // clicca sul tasto "No" per rimanere connessi
    } catch (error) {
       // bottone non apparso, ok...
    }
}

async function handleEmail(email) {
    // handle email reuse
    if(email == null) {
        if(fs.existsSync('./config.json')) {
            var data = fs.readFileSync('./config.json');
            try {
              let myObj = JSON.parse(data);
              email = myObj.email;
              console.log('Reusing previously saved email/username!\nIf you need to change it, use the -u argument')
            }
            catch (err) {
              term.red('There has been an error parsing your informations. Continuing in the manual way...\n')
              email = await promptQuestion("Email/username not saved. Please enter your email/username, MStreamDownloader will not ask for it next time: ");
              saveConfig({ email: email })
            }
        } else {
            email = await promptQuestion("Email/username not saved. Please enter your email/username, MStreamDownloader will not ask for it next time: ");
            saveConfig({ email: email })
        }
    }
    else {
        saveConfig({ email: email })
    }
    return email;
}

function saveConfig(infos) {
    var data = JSON.stringify(infos);
    try {
        fs.writeFileSync('./config.json', data);
        term.green('Email/username saved successfully. Next time you can avoid to insert it again.\n')
    } catch (e) {
        term.red('There has been an error saving your email/username offline. Continuing...\n');
    }
}

function doRequest(options) {
  return new Promise(function (resolve, reject) {
    request(options, function (error, res, body) {
      if (!error && (res.statusCode == 200 || res.statusCode == 403)) {
        resolve(body);
      } else {
        reject(error);
      }
    });
  });
}

function promptResChoice(question, count) {
  const readline = require('readline');
  const rl = readline.createInterface({
    input: process.stdin,
     output: process.stdout
  });

  return new Promise(function(resolve, reject) {
    var ask = function() {
      rl.question(question, function(answer) {
          if (!isNaN(answer) && parseInt(answer) < count && parseInt(answer) >= 0) {
            resolve(parseInt(answer), reject);
            rl.close();
          } else {
            console.log("\n* Wrong * - Please enter a number between 0 and " + (count-1) + "\n");
            ask();
        }
      });
    };
    ask();
  });
}

async function promptPassword(question) {
    return await promptQuestion(question, true)
}

async function promptQuestion(question) {
    return await promptQuestion(question, false)
}

function promptQuestion(question, hidden) {
  const readline = require('readline');
  const rl = readline.createInterface({
     input: process.stdin,
     output: process.stdout
  });

  if(hidden == true) {
      const stdin = process.openStdin();
      var onDataHandler = function(char) {
        char = char + '';
        switch (char) {
          case '\n':
          case '\r':
          case '\u0004':
            stdin.removeListener("data",onDataHandler);
            break;
          default:
            process.stdout.clearLine();
            readline.cursorTo(process.stdout, 0);
            process.stdout.write(question + Array(rl.line.length + 1).join('*'));
            break;
        }
      }
      process.stdin.on("data", onDataHandler);
  }

  return new Promise(function(resolve, reject) {
    var ask = function() {
      rl.question(question, function(answer) {
            if(hidden == true) rl.history = rl.history.slice(1);
            resolve(answer, reject);
            rl.close();
      });
    };
    ask();
  });
}


function rmDir(dir, rmSelf) {
    var files;
    rmSelf = (rmSelf === undefined) ? true : rmSelf;
    dir = dir + "/";
    try { files = fs.readdirSync(dir); } catch (e) { console.log("!Oops, directory not exist."); return; }
    if (files.length > 0) {
        files.forEach(function(x, i) {
            if (fs.statSync(dir + x).isDirectory()) {
                rmDir(dir + x);
            } else {
                fs.unlinkSync(dir + x);
            }
        });
    }
    if (rmSelf) {
        // check if user want to delete the directory or just the files in this directory
        fs.rmdirSync(dir);
    }
}


function sleep(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
}

async function extractCookies(page) {
    var jar = await page.cookies("https://.api.microsoftstream.com");
    var authzCookie = jar.filter(c => c.name === 'Authorization_Api')[0];
    var sigCookie = jar.filter(c => c.name === 'Signature_Api')[0];
    if (authzCookie == null || sigCookie == null) {
        await sleep(5000);
        var jar = await page.cookies("https://.api.microsoftstream.com");
        var authzCookie = jar.filter(c => c.name === 'Authorization_Api')[0];
        var sigCookie = jar.filter(c => c.name === 'Signature_Api')[0];
    }
    if (authzCookie == null || sigCookie == null) {
        console.error('Unable to read cookies. Try launching one more time, this is not an exact science.');
        process.exit(88);
    }
    return `Authorization=${authzCookie.value}; Signature=${sigCookie.value}`;
}

term.brightBlue(`Project originally based on https://github.com/snobu/destreamer\nFork powered by @sup3rgiu\nImprovements: - Multithreading download (much faster) - API Usage - Video Quality Choice - And More :)\n`);
sanityChecks();
let email = typeof argv.username === 'undefined' ? null : argv.username
let psw = typeof argv.password === 'undefined' ? null : argv.password
downloadVideo(argv.videoUrls, email, psw, argv.outputDirectory);
