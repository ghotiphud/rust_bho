#![allow(dead_code, 
    non_camel_case_types, 
    non_snake_case, 
    non_upper_case_globals, 
    unused_variables, 
    unused_imports)]

#[macro_use]
extern crate winapi;
extern crate user32;

use winapi::{GUID, REFGUID, REFIID, LPVOID, REFCLSID, HRESULT, BOOL, S_FALSE, S_OK,
    NOERROR, CLASS_E_CLASSNOTAVAILABLE};
use std::ptr;
use std::sync::atomic::{AtomicUsize, Ordering, ATOMIC_USIZE_INIT};


use std::ffi::OsStr;
use std::io::Error;
use std::iter::once;
use std::os::windows::ffi::OsStrExt;

#[link(name = "ole32")]
extern "system" {
    pub fn IsEqualGUID(rguid1: REFGUID, rguid2: REFGUID) -> BOOL;
    //pub fn IsEqualCLSID(rclsid1: REFCLSID, rclsid2: REFCLSID) -> BOOL;
    //fn IsEqualIID(riid1: REFGUID, riid2: REFGUID) -> BOOL;
}

macro_rules! EXPORT_GUID {
    (
        $name:ident, $l:expr, $w1:expr, $w2:expr,
        $b1:expr, $b2:expr, $b3:expr, $b4:expr, $b5:expr, $b6:expr, $b7:expr, $b8:expr
    ) => {
        #[no_mangle]
        pub static $name: GUID = GUID {
            Data1: $l,
            Data2: $w1,
            Data3: $w2,
            Data4: [$b1, $b2, $b3, $b4, $b5, $b6, $b7, $b8],
        };
    }
}

// COM Interface GUIDs
EXPORT_GUID!{IID_IUnknown, 0x00000000, 0x0000, 0x0000, 0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}
EXPORT_GUID!{IID_IClassFactory, 0x00000001, 0x0000, 0x0000, 0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}
EXPORT_GUID!{IID_IObjectWithSite, 0xFC4801A3, 0x2BA9, 0x11CF, 0xA2, 0x29, 0x00, 0xAA, 0x00, 0x3D, 0x73, 0x52}

// Custom GUIDs
// {C1626E91-F628-4F69-94ED-4DB4564ECDF5}
EXPORT_GUID!(CLSID_IEExtension, 0xc1626e91, 0xf628, 0x4f69, 0x94, 0xed, 0x4d, 0xb4, 0x56, 0x4e, 0xcd, 0xf5);
// {1B474362-E889-4A60-BB89-AC457A986A97}
EXPORT_GUID!(IID_IEExtension, 0x1b474362, 0xe889, 0x4a60, 0xbb, 0x89, 0xac, 0x45, 0x7a, 0x98, 0x6a, 0x97);


#[cfg(not(feature = "rusty"))]
mod cstyle;
#[cfg(not(feature = "rusty"))]
pub use cstyle::{IEExtension, IEExtensionClassFactory};
#[cfg(not(feature = "rusty"))]
use cstyle::{IEExtensionClassFactory_STATIC};

// #[cfg(feature = "rusty")]
// mod ruststyle;
// #[cfg(feature = "rusty")]
// pub use ruststyle::*;

// COM Entrypoint fns
// HRESULT __stdcall DllGetClassObject(
//   _In_  REFCLSID rclsid,
//   _In_  REFIID   riid,libloading = "0.3"
//   _Out_ LPVOID   *ppv
// );
#[no_mangle]
pub unsafe extern "system" fn DllGetClassObject(rclsid: REFGUID, riid: REFIID, ppv: *mut LPVOID) -> HRESULT {
    println!("DllGetClassObject");
    let hr: HRESULT;

    if IsEqualGUID(rclsid, &CLSID_IEExtension) != 0 {
        hr = (*IEExtensionClassFactory_STATIC).QueryInterface(riid, ppv);
    } else {
        *ppv = ptr::null_mut();
        hr = CLASS_E_CLASSNOTAVAILABLE;
    }

    hr
}

static OUTSTANDING_OBJECTS: AtomicUsize = ATOMIC_USIZE_INIT;
static LOCK_COUNT: AtomicUsize = ATOMIC_USIZE_INIT;

// HRESULT __stdcall DllCanUnloadNow(void);
#[no_mangle]
pub unsafe extern "system" fn DllCanUnloadNow() -> HRESULT {
    println!("DllCanUnloadNow");
    let obj_count = OUTSTANDING_OBJECTS.load(Ordering::SeqCst);
    let lock_count = LOCK_COUNT.load(Ordering::SeqCst);
    if (obj_count | lock_count) == 0 { S_FALSE } else { S_OK }
}

// HRESULT __stdcall DllRegisterServer(void);
#[no_mangle]
pub unsafe extern "system" fn DllRegisterServer() -> HRESULT {
    println!("DllRegisterServer");
    NOERROR
}

// HRESULT __stdcall DllUnregisterServer(void);
#[no_mangle]
pub unsafe extern "system" fn DllUnregisterServer() -> HRESULT {
    println!("DllUnregisterServer");
    NOERROR
}

#[cfg(test)]
mod tests {
    use super::{IEExtension};
    use std::ptr;

    #[test]
    fn initialize_IEExtension() {
        let mut ext = IEExtension::new(1);
        let mut ext2 = IEExtension::new(1);

        unsafe{
            ext.SetSite(&mut *ext2);
        }
    }
}