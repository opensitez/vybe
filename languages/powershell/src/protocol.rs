//! PowerShell protocol slot mapping for class normalization.

use vybe_ast::class_normalize::types::SpecialMethodKind;

/// Map a PowerShell method name to a canonical runtime role.
pub fn canonical_method(name: &str) -> (String, Option<SpecialMethodKind>) {
    use SpecialMethodKind::*;

    if name.starts_with('~') {
        return ("destructor".into(), Some(Destructor));
    }

    let canonical = name.to_ascii_lowercase();
    match canonical.as_str() {
        "tostring" => ("tostring".into(), Some(ToString)),
        "gethashcode" => ("hash".into(), Some(Hash)),
        "equals" => ("eq".into(), Some(Eq)),
        "compareto" => ("compare".into(), Some(Compare)),
        "getenumerator" => ("iterator".into(), Some(Iterator)),
        "movenext" => ("next".into(), Some(Next)),
        "dispose" => ("exit".into(), Some(Exit)),
        "finalize" => ("destructor".into(), Some(Destructor)),
        "clone" => ("clone".into(), Some(Clone)),

        // Indexer-like roles.
        "getitem" => ("getitem".into(), Some(GetItem)),
        "offsetget" => ("getitem".into(), Some(GetItem)),
        "setitem" => ("setitem".into(), Some(SetItem)),
        "offsetset" => ("setitem".into(), Some(SetItem)),
        "hasitem" => ("hasitem".into(), Some(HasItem)),
        "offsetexists" => ("hasitem".into(), Some(HasItem)),
        "delitem" => ("delitem".into(), Some(DelItem)),
        "offsetunset" => ("delitem".into(), Some(DelItem)),

        // Reflection-like roles.
        "getattr" => ("getattr".into(), Some(GetAttr)),
        "setattr" => ("setattr".into(), Some(SetAttr)),
        "hasattr" => ("hasattr".into(), Some(HasAttr)),
        "delattr" => ("delattr".into(), Some(DelAttr)),

        // Callable / missing handlers.
        "call" => ("call".into(), Some(Call)),
        "__invoke" => ("call".into(), Some(Call)),
        "callmissing" => ("callmissing".into(), Some(CallMissing)),
        "__call" => ("callmissing".into(), Some(CallMissing)),
        "callstatic" => ("callstatic".into(), Some(CallStatic)),
        "__callstatic" => ("callstatic".into(), Some(CallStatic)),

        // Context manager hooks.
        "enter" => ("enter".into(), Some(Enter)),
        "exit" => ("exit".into(), Some(Exit)),

        "contains" => ("contains".into(), Some(Contains)),
        _ => (canonical, None),
    }
}
