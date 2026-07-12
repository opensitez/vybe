//! WinForms designer form-module registry now lives in the dotnet platform
//! (`vybe_platform_dotnet::winforms::form_modules`); VB / C# register their own
//! modules through it. Re-exported here so existing
//! `crate::projects::form_modules::…` paths keep resolving.

pub use vybe_platform_dotnet::winforms::form_modules::*;
