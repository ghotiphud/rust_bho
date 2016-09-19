# rust_bho
Proof of Concept Internet Explorer Browser Helper Object (BHO) implemented in Rust

## Usage:
I've only tested this when compiled as i686, so either modify this for 64-bit or `rustup override set stable-i686-pc-windows-msvc`.

To build: `cargo build --release`

If you've ever dabbled in COM programming or BHOs you'll know that the windows registry is key to it all.

Minimal Registry data: (put in file RegBHO.reg, replace `{path_to_rust_bho}`, then double click)
```
Windows Registry Editor Version 5.00

[HKEY_CLASSES_ROOT\WOW6432Node\CLSID\{C1626E91-F628-4F69-94ED-4DB4564ECDF5}]
@="rust_bho"

[HKEY_CLASSES_ROOT\WOW6432Node\CLSID\{C1626E91-F628-4F69-94ED-4DB4564ECDF5}\InprocServer32]
@="{path_to_rust_bho}\\target\\release\\rust_bho.dll"
"ThreadingModel"="Apartment"

[HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Explorer\Browser Helper Objects\{C1626E91-F628-4F69-94ED-4DB4564ECDF5}]
@="rust_bho"
"NoExplorer"="1"
```

To test that the com object was registered correctly: `cargo run`

Start IE and you should see some message boxes.  Remove the registry keys to disable the whole thing.

Credit to: 
http://www.codeproject.com/Articles/13601/COM-in-plain-C
and
http://www.codeproject.com/Articles/37044/Writing-a-BHO-in-Plain-C
