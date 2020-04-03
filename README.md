# MStreamDownloader

## Download Microsoft Stream videos.

This project was originally based on https://github.com/snobu/destreamer

Improvements in this fork:
 - Multithreading download through aria2c (much faster)
 - Possibility to choose the video resolution
 - API Usage
 - Much More :)

**PoliMi students can use this fork too: https://github.com/sup3rgiu/PoliDown**


## PREREQS

* [**Node.js**](https://nodejs.org/it/download/): anything above v8.0 seems to work.
* [**aria2**](https://github.com/aria2/aria2/releases): this needs to be in your `$PATH` (for example, copy aria2c.exe to c:\windows). downloader calls `aria2c` with a bunch of arguments in order to improve the download speed.
* [**ffmpeg**](https://www.ffmpeg.org/download.html): a recent version (year 2019 or above), in [`$PATH`](https://www.thewindowsclub.com/how-to-install-ffmpeg-on-windows-10).

## Windows Installation instructions
(On others OS is pretty much the same)

**N.B: this video was done for https://github.com/sup3rgiu/PoliDown, but the process is the same.**

[![https://user-images.githubusercontent.com/7725068/76635047-21a89080-6547-11ea-8da9-31831ca7620a.png](https://user-images.githubusercontent.com/7725068/76635345-a1cef600-6547-11ea-991b-d115946ed556.png)](http://www.youtube.com/watch?v=iZgea4t5YW4 "downloader Windows Installation Instructions")


## USAGE

* Clone this repo
* `cd` into the cloned folder
* `npm install` to install dependencies

Default usage:
```
$ node downloader.js -v "https://web.microsoftstream.com/video/VIDEO-1"
```

Show options:
```
$ node downloader.js -h

Options:
  --version              Show version number                           [boolean]
  -v, --videoUrls                                             [array] [required]
  -u, --username         Your Microsoft Email                           [string]
  -m, --polimi           PoliMi Login. If set, use Codice Persona as username
                                                    [boolean] [default: false]
  -o, --outputDirectory                             [string] [default: "videos"]
  -q, --quality          Video Quality, usually [0-5]                   [number]
  -k, --noKeyring        Do not use system keyring    [boolean] [default: false]
  -c, --conn             Number of simultaneous connections [1-16]
                                                        [number] [default: 16]
  -h, --help             Show help                                     [boolean]
```

Multiple videos download:
```
$ node downloader.js -v "https://web.microsoftstream.com/video/VIDEO-1" "https://web.microsoftstream.com/video/VIDEO-2" "https://web.microsoftstream.com/video/VIDEO-3"
```

Define default video quality [0-5] (avoid manual prompt for each video):
```
$ node downloader.js -v "https://web.microsoftstream.com/video/VIDEO-1" -q 4
```

Output directory (relative or absoulte path):
```
$ node downloader.js -v "https://web.microsoftstream.com/video/VIDEO-1" -o "/my/path/here"
```

Define number of simultaneous connection used to download [1-16, default: 16]:
```
$ node downloader.js -v "https://web.microsoftstream.com/video/VIDEO-1" -c 10
```

Do not use system keyring to save the password:
```
$ node downloader.js -v "https://web.microsoftstream.com/video/VIDEO-1" -k
```

Use Politecnico di Milano (Italy) login instead of default one (only for PoliMi students!)
(PoliMi students can use this fork too: https://github.com/sup3rgiu/PoliDown)
```
$ node downloader.js -v "https://web.microsoftstream.com/video/VIDEO-1" -m
```


You can omit the password argument. MStreamDownloader will ask for it interactively and then save it securely in system's keychain for the next use.

## EXPECTED OUTPUT

```
Project originally based on https://github.com/snobu/destreamer
Fork powered by @sup3rgiu
Improvements: - Multithreading download (much faster) - API Usage - Video Quality Choice - And More :)
Using aria2 version 1.35.0
Using ffmpeg version git-2020-03-06-cfd9a65 Copyright (c) 2000-2020 the FFmpeg developers

Launching headless Chrome to perform the OpenID Connect dance...
Navigating to STS login page...
We are logged in.
Got required authentication cookies.

Start downloading video: https://web.microsoftstream.com/video/d1e6c909-3189-488f-8172-e88947249f02

Video title is: SALIONI ALBERTO  088805-FISICA TECNICA (712804)

[0] 320x180
[1] 480x270
[2] 640x360
[3] 960x540
[4] 1280x720
[5] 1920x1080
Choose the desired resolution: 5

03/10 20:47:14 [NOTICE] Downloading 898 item(s)

[...]

At this point Chrome's job is done, shutting it down...
Done!
```

The video is now saved under `videos/`, or whatever the `outputDirectory` argument points to.
