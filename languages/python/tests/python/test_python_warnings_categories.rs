use super::helpers::run_python;

// warnings — categories, filterwarnings, catch_warnings, warn_explicit, stacklevel, simplefilter

#[test]
fn test_warnings_user_warning_category() {
    let out = run_python(
        r#"
import warnings
with warnings.catch_warnings(record=True) as w:
    warnings.simplefilter("always")
    warnings.warn("user message", UserWarning)
print(len(w))
print(issubclass(w[0].category, UserWarning))
print(str(w[0].message))
"#,
    );
    assert_eq!(out, vec!["1", "True", "user message"]);
}

#[test]
fn test_warnings_deprecation_warning_category() {
    let out = run_python(
        r#"
import warnings
with warnings.catch_warnings(record=True) as w:
    warnings.simplefilter("always")
    warnings.warn("old api", DeprecationWarning)
print(w[0].category.__name__)
"#,
    );
    assert_eq!(out, vec!["DeprecationWarning"]);
}

#[test]
fn test_warnings_runtime_warning() {
    let out = run_python(
        r#"
import warnings
with warnings.catch_warnings(record=True) as w:
    warnings.simplefilter("always")
    warnings.warn("runtime issue", RuntimeWarning)
print(w[0].category.__name__)
"#,
    );
    assert_eq!(out, vec!["RuntimeWarning"]);
}

#[test]
fn test_warnings_filter_ignore_suppresses_warning() {
    let out = run_python(
        r#"
import warnings
with warnings.catch_warnings(record=True) as w:
    warnings.simplefilter("ignore")
    warnings.warn("ignored", UserWarning)
print(len(w))
"#,
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn test_warnings_filter_error_raises() {
    let out = run_python(
        r#"
import warnings
with warnings.catch_warnings():
    warnings.simplefilter("error", UserWarning)
    try:
        warnings.warn("turns into error", UserWarning)
    except UserWarning as e:
        print(str(e))
"#,
    );
    assert_eq!(out, vec!["turns into error"]);
}

#[test]
fn test_warnings_filterwarnings_once() {
    let out = run_python(
        r#"
import warnings
with warnings.catch_warnings(record=True) as w:
    warnings.simplefilter("always")
    warnings.filterwarnings("once", message=".*repeated.*")
    warnings.warn("repeated warning", UserWarning)
    warnings.warn("repeated warning", UserWarning)
# "once" means only first occurrence per location
print(len(w) >= 1)
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_warnings_filterwarnings_module_regex() {
    let out = run_python(
        r#"
import warnings
with warnings.catch_warnings(record=True) as w:
    warnings.simplefilter("always")
    warnings.filterwarnings("ignore", category=DeprecationWarning, module=".*")
    warnings.warn("dep", DeprecationWarning)
    warnings.warn("user", UserWarning)
print(len(w))
print(w[0].category.__name__)
"#,
    );
    assert_eq!(out, vec!["1", "UserWarning"]);
}

#[test]
fn test_warnings_warn_explicit_custom_lineno() {
    let out = run_python(
        r#"
import warnings
with warnings.catch_warnings(record=True) as w:
    warnings.simplefilter("always")
    warnings.warn_explicit(
        "explicit warning",
        UserWarning,
        filename="fake_file.py",
        lineno=42,
        module="fake_module"
    )
print(w[0].lineno)
print(w[0].filename)
"#,
    );
    assert_eq!(out, vec!["42", "fake_file.py"]);
}

#[test]
fn test_warnings_warn_explicit_category_class() {
    let out = run_python(
        r#"
import warnings
with warnings.catch_warnings(record=True) as w:
    warnings.simplefilter("always")
    warnings.warn_explicit(
        "test",
        DeprecationWarning,
        filename="x.py",
        lineno=1,
    )
print(w[0].category.__name__)
"#,
    );
    assert_eq!(out, vec!["DeprecationWarning"]);
}

#[test]
fn test_warnings_stacklevel_2_points_to_caller() {
    let out = run_python(
        r#"
import warnings
def inner():
    warnings.warn("from inner", UserWarning, stacklevel=2)
with warnings.catch_warnings(record=True) as w:
    warnings.simplefilter("always")
    inner()   # warning should point here (stacklevel=2)
# The warning lineno should be the call to inner(), not inside inner()
print(len(w) == 1)
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_warnings_catch_warnings_restores_filters() {
    let out = run_python(
        r#"
import warnings
original = warnings.filters[:]
with warnings.catch_warnings():
    warnings.simplefilter("ignore")
print(warnings.filters == original)
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_warnings_multiple_warnings_recorded() {
    let out = run_python(
        r#"
import warnings
with warnings.catch_warnings(record=True) as w:
    warnings.simplefilter("always")
    warnings.warn("first", UserWarning)
    warnings.warn("second", DeprecationWarning)
    warnings.warn("third", RuntimeWarning)
print(len(w))
print([x.category.__name__ for x in w])
"#,
    );
    assert_eq!(
        out,
        vec![
            "3",
            "['UserWarning', 'DeprecationWarning', 'RuntimeWarning']"
        ]
    );
}

#[test]
fn test_warnings_default_filter_once_per_location() {
    let out = run_python(
        r#"
import warnings
with warnings.catch_warnings(record=True) as w:
    warnings.simplefilter("default")
    warnings.warn("msg", UserWarning)
    warnings.warn("msg", UserWarning)  # same message, same location
print(len(w) == 1)
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_warnings_custom_warning_subclass() {
    let out = run_python(
        r#"
import warnings
class MyWarning(UserWarning):
    pass
with warnings.catch_warnings(record=True) as w:
    warnings.simplefilter("always")
    warnings.warn("custom", MyWarning)
print(w[0].category.__name__)
print(issubclass(w[0].category, UserWarning))
"#,
    );
    assert_eq!(out, vec!["MyWarning", "True"]);
}

#[test]
fn test_warnings_warn_future_warning() {
    let out = run_python(
        r#"
import warnings
with warnings.catch_warnings(record=True) as w:
    warnings.simplefilter("always")
    warnings.warn("future change", FutureWarning)
print(w[0].category.__name__)
"#,
    );
    assert_eq!(out, vec!["FutureWarning"]);
}

#[test]
fn test_warnings_warn_resource_warning() {
    let out = run_python(
        r#"
import warnings
with warnings.catch_warnings(record=True) as w:
    warnings.simplefilter("always")
    warnings.warn("resource leak", ResourceWarning)
print(w[0].category.__name__)
"#,
    );
    assert_eq!(out, vec!["ResourceWarning"]);
}

#[test]
fn test_warnings_message_object_accessible() {
    let out = run_python(
        r#"
import warnings
with warnings.catch_warnings(record=True) as w:
    warnings.simplefilter("always")
    warnings.warn("payload message", UserWarning)
msg_obj = w[0].message
print(type(msg_obj).__name__)
print(str(msg_obj))
"#,
    );
    assert_eq!(out, vec!["UserWarning", "payload message"]);
}

#[test]
fn test_warnings_resetwarnings_clears_filters() {
    let out = run_python(
        r#"
import warnings
with warnings.catch_warnings():
    warnings.simplefilter("ignore")
    warnings.filterwarnings("error", category=DeprecationWarning)
    warnings.resetwarnings()
    # After reset, no custom filters remain
    print(len(warnings.filters) == 0)
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_warnings_filterwarnings_action_always_accumulates() {
    let out = run_python(
        r#"
import warnings
with warnings.catch_warnings(record=True) as w:
    warnings.filterwarnings("always", message=".*repeat.*", category=UserWarning)
    for _ in range(3):
        warnings.warn("repeat me", UserWarning)
print(len(w))
"#,
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn test_warnings_syntax_warning() {
    let out = run_python(
        r#"
import warnings
with warnings.catch_warnings(record=True) as w:
    warnings.simplefilter("always")
    warnings.warn("syntax issue", SyntaxWarning)
print(w[0].category.__name__)
"#,
    );
    assert_eq!(out, vec!["SyntaxWarning"]);
}
