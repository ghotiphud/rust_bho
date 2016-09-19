#![allow(non_snake_case)]

extern crate rust_bho;
//extern crate libloading as lib;
extern crate winapi;

use rust_bho::*;
use winapi::{GUID, HRESULT, REFGUID, REFIID, LPVOID,
REFCLSID, IUnknown, DWORD, c_void,
COINIT_APARTMENTTHREADED, CLSCTX_INPROC_SERVER};

use std::ptr;

#[link(name = "ole32")]
extern "stdcall" {
    fn CoCreateInstance(rclsid: REFCLSID, pUnkOuter: *mut IUnknown, dwClsContext: DWORD, riid: REFIID, ppv: *mut *mut c_void) -> HRESULT;
    fn CoInitializeEx(pvReserved: *mut c_void, dwCoInit: DWORD) -> HRESULT;
    fn CoUninitialize();
}

pub fn main() {
    // To see exports from dll
    // VS Developer Console> dumpbin /exports rust_bho.dll

    unsafe{
        println!("{:?}", 
            CoInitializeEx(ptr::null_mut(), COINIT_APARTMENTTHREADED) );

        let mut ie_obj_ptr = ptr::null_mut();

        println!("{:?}", 
            CoCreateInstance(&CLSID_IEExtension, ptr::null_mut(), CLSCTX_INPROC_SERVER, &IID_IObjectWithSite, &mut ie_obj_ptr));

        let mut ie_obj: &mut IEExtension = std::mem::transmute(ie_obj_ptr);

        let mut ext = IEExtension::new(1);
        ie_obj.SetSite(&mut *ext);
    }

    // let rust_bho = lib::Library::new("target/debug/rust_bho.dll").expect("library not loaded");

    // unsafe {
    //     let CLSID_IEExtension: lib::Symbol<*const GUID> =
    //         rust_bho.get(b"CLSID_IEExtension\0").unwrap();
    //     let IID_IUnknown: lib::Symbol<*const GUID> =
    //         rust_bho.get(b"IID_IUnknown\0").unwrap();

    //     let DllCanUnloadNow: lib::Symbol<unsafe extern "system" fn() -> HRESULT> = 
    //         rust_bho.get(b"DllCanUnloadNow\0").unwrap();
    //     let DllGetClassObject: lib::Symbol<unsafe extern "system" fn(REFGUID, REFIID, *mut LPVOID) -> HRESULT> = 
    //         rust_bho.get(b"DllGetClassObject\0").unwrap();

    //     println!("{:?}", **CLSID_IEExtension);
    //     println!("{:?}", **IID_IUnknown);


    //     println!("{:?}", DllCanUnloadNow() != 0);

    //     let mut cls_factory_ptr = ptr::null_mut();
    //     let result = DllGetClassObject(*CLSID_IEExtension, *IID_IUnknown, &mut cls_factory_ptr);
    //     println!("{:?}", result);
    //     println!("{:?}", cls_factory_ptr);

    // }

    // unsafe {
    //     println!("{:?}", DllCanUnloadNow() != 0);

    //     let mut cls_factory_ptr = ptr::null_mut();
    //     let result = DllGetClassObject(&CLSID_IEExtension, &IID_IUnknown, &mut cls_factory_ptr);
    //     println!("{:?}", result);
    //     println!("{:?}", cls_factory_ptr);

    //     let cls_factory: &IEExtensionClassFactory = std::mem::transmute(cls_factory_ptr);

    //     let mut com_obj_ptr = ptr::null_mut();

    //     println!("{:?}", cls_factory as (*const _));

    //     let result2 = cls_factory.CreateInstance(ptr::null_mut(), &IID_IObjectWithSite, &mut com_obj_ptr);

    //     println!("{:?}", result2);
    //     println!("{:?}", DllCanUnloadNow() != 0);
    //     println!("{:?}", com_obj_ptr);

    //     let mut com_obj: &mut IEExtension = std::mem::transmute(com_obj_ptr);
    //     println!("{:?}", com_obj as (*const _));

    //     let mut ext = IEExtension::new(1);
    //     com_obj.SetSite(&mut *ext);
    // }
}