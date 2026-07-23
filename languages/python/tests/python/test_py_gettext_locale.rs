use super::helpers::run_python;

// ═══════════════════════════════════════════════════════════
// Python: locale + gettext — internationalization, localization, currency/number formatting
// ═══════════════════════════════════════════════════════════

#[test]
fn test_py_locale_getlocale_setlocale() {
    let src = r#"
import locale

current = locale.getlocale()
print(isinstance(current, tuple))

try:
    locale.setlocale(locale.LC_ALL, "C")
    print(locale.getlocale())
except locale.Error:
    print("Locale setting failed")
"#;
    assert_eq!(run_python(src), vec!["True", "(None, None)"]);
}

#[test]
fn test_py_locale_format_string() {
    let src = r#"
import locale

locale.setlocale(locale.LC_ALL, "C")
formatted = locale.format_string("%d", 1234567, grouping=True)
print(formatted)
"#;
    assert_eq!(run_python(src), vec!["1234567"]);
}

#[test]
fn test_py_locale_strcoll_strxfrm() {
    let src = r#"
import locale

locale.setlocale(locale.LC_ALL, "C")
print(locale.strcoll("apple", "banana") < 0)
print(locale.strcoll("zebra", "apple") > 0)
print(isinstance(locale.strxfrm("test"), str))
"#;
    assert_eq!(run_python(src), vec!["True", "True", "True"]);
}

#[test]
fn test_py_gettext_null_translation() {
    let src = r#"
import gettext

t = gettext.NullTranslations()
print(t.gettext("Hello World"))
print(t.ngettext("apple", "apples", 1))
print(t.ngettext("apple", "apples", 3))
"#;
    assert_eq!(run_python(src), vec!["Hello World", "apple", "apples"]);
}

#[test]
fn test_py_gettext_install_global() {
    let src = r#"
import gettext

t = gettext.NullTranslations()
t.install()

# Global _ function is installed
print(_("Translatable string"))
"#;
    assert_eq!(run_python(src), vec!["Translatable string"]);
}

#[test]
fn test_py_gettext_gnu_translations_dict() {
    let src = r#"
import gettext, io

catalog = {
    ("Hello World", None): "Bonjour le monde",
    ("Goodbye", None): "Au revoir",
}

class DictTranslation(gettext.NullTranslations):
    def __init__(self, mapping):
        super().__init__()
        self._mapping = mapping

    def gettext(self, message):
        return self._mapping.get((message, None), message)

trans = DictTranslation(catalog)
print(trans.gettext("Hello World"))
print(trans.gettext("Goodbye"))
print(trans.gettext("Unknown"))
"#;
    assert_eq!(
        run_python(src),
        vec!["Bonjour le monde", "Au revoir", "Unknown"]
    );
}

#[test]
fn test_py_locale_conv_currency() {
    let src = r#"
import locale

locale.setlocale(locale.LC_ALL, "C")
conv = locale.localeconv()
print(isinstance(conv, dict))
print("currency_symbol" in conv)
print("decimal_point" in conv)
"#;
    assert_eq!(run_python(src), vec!["True", "True", "True"]);
}
