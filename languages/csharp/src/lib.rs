pub mod forms;
pub mod normalize_class;
pub mod walker;

use pest_derive::Parser;

#[derive(Parser)]
#[grammar = "src/grammar.pest"]
pub(crate) struct CSharpParser;

/// Parse C# source into the common AST.
pub fn parse(source: &str) -> Result<vybe_ast::Module, String> {
    walker::parse(source)
}

/// Embedded profile TOML source.
pub fn profile_source() -> &'static str {
    include_str!("profile")
}

/// Register this language with the shared plugin registry (dylib entry point).
pub fn register() {
    vybe_plugin::registry::register_language(vybe_plugin::registry::LanguagePlugin {
        name: "csharp",
        parse,
        profile_source,
        emit_dispatch: None,
        normalize_class: Some(normalize_class::normalize_class),
        register_tree: None,
    });
    vybe_platform_dotnet::winforms::form_modules::register(
        vybe_platform_dotnet::winforms::form_modules::FormModuleLanguage {
            name: "csharp",
            load_designer: forms::load_designer,
            save_designer: forms::save_designer,
            generate_designer_code: forms::generate_designer_code,
            generate_user_code_stub: forms::generate_user_code_stub,
        },
    );
}

/// This crate as a [`vybe_plugin::Plugin`] — its `init` registers the
/// language (and any forms) with the shared framework. Also the dylib entry point.
pub struct Plugin;
impl vybe_plugin::Plugin for Plugin {
    fn name(&self) -> &'static str {
        "csharp"
    }
    fn init(&self, _fw: &mut vybe_plugin::Framework<'_>) {
        register();
    }
}
