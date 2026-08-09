use super::helpers::run_python;

// gettext — GNUTranslations, NullTranslations, gettext, ngettext, pgettext, npgettext, bindtextdomain, textdomain, translation, dgettext, dngettext

#[test]
fn test_gettext_null_translations_fallback() {
    let out = run_python(
        r#"
import gettext
t = gettext.NullTranslations()
print(t.gettext("Hello"))
print(t.ngettext("Apple", "Apples", 1))
print(t.ngettext("Apple", "Apples", 5))
"#,
    );
    assert_eq!(out, vec!["Hello", "Apple", "Apples"]);
}

#[test]
fn test_gettext_null_translations_pgettext_context() {
    let out = run_python(
        r#"
import gettext
t = gettext.NullTranslations()
print(t.pgettext("menu", "File"))
print(t.npgettext("menu", "File", "Files", 1))
print(t.npgettext("menu", "File", "Files", 2))
"#,
    );
    assert_eq!(out, vec!["File", "File", "Files"]);
}

#[test]
fn test_gettext_bindtextdomain_and_textdomain() {
    let out = run_python(
        r#"
import gettext
d = gettext.bindtextdomain("myapp", "/usr/share/locale")
dom = gettext.textdomain("myapp")
print(d)
print(dom)
"#,
    );
    assert_eq!(out, vec!["/usr/share/locale", "myapp"]);
}

#[test]
fn test_gettext_translation_fallback_true() {
    let out = run_python(
        r#"
import gettext
trans = gettext.translation("non_existent_domain", fallback=True)
print(isinstance(trans, gettext.NullTranslations))
print(trans.gettext("Test"))
"#,
    );
    assert_eq!(out, vec!["True", "Test"]);
}

#[test]
fn test_gettext_translation_fallback_false_raises_file_not_found() {
    let out = run_python(
        r#"
import gettext
try:
    gettext.translation("non_existent_domain", fallback=False)
except OSError:
    print("OSError")
"#,
    );
    assert_eq!(out, vec!["OSError"]);
}

#[test]
fn test_gettext_gnu_translations_dict_simulation() {
    let out = run_python(
        r#"
import gettext, io

gt = gettext.GNUTranslations()
gt._catalog = {
    "Hello": "Bonjour",
    ("Apple", 1): "Pomme",
    ("Apple", 2): "Pommes"
}
print(gt.gettext("Hello"))
print(gt.gettext("Missing"))
"#,
    );
    assert_eq!(out, vec!["Bonjour", "Missing"]);
}

#[test]
fn test_gettext_install_builtins_underscore() {
    let out = run_python(
        r#"
import gettext, builtins
t = gettext.NullTranslations()
t.install()
print(hasattr(builtins, "_"))
print(_("Installed test"))
"#,
    );
    assert_eq!(out, vec!["True", "Installed test"]);
}

#[test]
fn test_gettext_dgettext_and_dngettext() {
    let out = run_python(
        r#"
import gettext
print(gettext.dgettext("domain", "Hello"))
print(gettext.dngettext("domain", "Item", "Items", 2))
"#,
    );
    assert_eq!(out, vec!["Hello", "Items"]);
}

#[test]
fn test_gettext_pgettext_separator_handling() {
    let out = run_python(
        r#"
import gettext
res = gettext.pgettext("context", "msgid")
print(res)
"#,
    );
    assert_eq!(out, vec!["msgid"]);
}

#[test]
fn test_gettext_find_localedir_lookup() {
    let out = run_python(
        r#"
import gettext
res = gettext.find("domain", localedir=None, languages=["en"])
print(res is None)
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_gettext_null_translations_info() {
    let out = run_python(
        r#"
import gettext
t = gettext.NullTranslations()
print(t.info())
t._info = {"content-type": "text/plain"}
print(t.info())
"#,
    );
    assert_eq!(out, vec!["{}", "{'content-type': 'text/plain'}"]);
}

#[test]
fn test_gettext_null_translations_charset() {
    let out = run_python(
        r#"
import gettext
t = gettext.NullTranslations()
print(t.charset())
"#,
    );
    assert_eq!(out, vec!["None"]);
}

#[test]
fn test_gettext_lgettext_deprecated_or_attribute_check() {
    let out = run_python(
        r#"
import gettext
t = gettext.NullTranslations()
print(hasattr(t, "gettext"))
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_gettext_install_names_parameter() {
    let out = run_python(
        r#"
import gettext, builtins
t = gettext.NullTranslations()
t.install(names=["ngettext"])
print(hasattr(builtins, "ngettext"))
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_gettext_dpgettext_and_dnpgettext() {
    let out = run_python(
        r#"
import gettext
print(gettext.dpgettext("dom", "ctx", "msg"))
print(gettext.dnpgettext("dom", "ctx", "singular", "plural", 3))
"#,
    );
    assert_eq!(out, vec!["msg", "plural"]);
}

#[test]
fn test_gettext_gnu_translations_plural_evaluation() {
    let out = run_python(
        r#"
import gettext
gt = gettext.GNUTranslations()
gt.plural = lambda n: int(n != 1)
gt._catalog = {
    ("dog", 0): "chien",
    ("dog", 1): "chiens"
}
print(gt.ngettext("dog", "dogs", 1))
print(gt.ngettext("dog", "dogs", 2))
"#,
    );
    assert_eq!(out, vec!["chien", "chiens"]);
}

#[test]
fn test_gettext_output_charset() {
    let out = run_python(
        r#"
import gettext
t = gettext.NullTranslations()
print(hasattr(t, "output_charset"))
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_gettext_add_fallback_chain() {
    let out = run_python(
        r#"
import gettext
t1 = gettext.NullTranslations()
t2 = gettext.NullTranslations()
t1.add_fallback(t2)
print(t1._fallback is t2)
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_gettext_global_gettext_module_level_helpers() {
    let out = run_python(
        r#"
import gettext
print(gettext.gettext("test"))
print(gettext.ngettext("cat", "cats", 1))
"#,
    );
    assert_eq!(out, vec!["test", "cat"]);
}

#[test]
fn test_gettext_class_hierarchy() {
    let out = run_python(
        r#"
import gettext
print(issubclass(gettext.GNUTranslations, gettext.NullTranslations))
"#,
    );
    assert_eq!(out, vec!["True"]);
}
