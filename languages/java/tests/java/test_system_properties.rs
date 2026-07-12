use crate::helpers::run_main;

#[test]
fn properties_new_is_empty() {
    let out = run_main(
        r#"java.util.Properties p = new java.util.Properties(); System.out.println(p.isEmpty());"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn properties_set_and_get() {
    let out = run_main(
        r#"java.util.Properties p = new java.util.Properties(); p.setProperty("key", "value"); System.out.println(p.getProperty("key"));"#,
    );
    assert_eq!(out, vec!["value"]);
}

#[test]
fn properties_get_missing_returns_null() {
    let out = run_main(
        r#"java.util.Properties p = new java.util.Properties(); System.out.println(p.getProperty("missing") == null);"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn properties_get_with_default() {
    let out = run_main(
        r#"java.util.Properties p = new java.util.Properties(); System.out.println(p.getProperty("missing", "fallback"));"#,
    );
    assert_eq!(out, vec!["fallback"]);
}

#[test]
fn properties_set_property_returns_old() {
    let out = run_main(
        r#"java.util.Properties p = new java.util.Properties(); p.setProperty("k", "v1"); System.out.println(p.setProperty("k", "v2"));"#,
    );
    assert_eq!(out, vec!["v1"]);
}

#[test]
fn properties_contains_key() {
    let out = run_main(
        r#"java.util.Properties p = new java.util.Properties(); p.setProperty("a", "1"); System.out.println(p.containsKey("a"));"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn properties_contains_value() {
    let out = run_main(
        r#"java.util.Properties p = new java.util.Properties(); p.setProperty("a", "1"); System.out.println(p.containsValue("1"));"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn properties_remove_key() {
    let out = run_main(
        r#"java.util.Properties p = new java.util.Properties(); p.setProperty("a", "1"); p.remove("a"); System.out.println(p.isEmpty());"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn properties_size_after_puts() {
    let out = run_main(
        r#"java.util.Properties p = new java.util.Properties(); p.setProperty("a", "1"); p.setProperty("b", "2"); System.out.println(p.size());"#,
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn properties_string_property_names() {
    let out = run_main(
        r#"java.util.Properties p = new java.util.Properties(); p.setProperty("x", "1"); System.out.println(p.stringPropertyNames().contains("x"));"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn properties_put_overwrites() {
    let out = run_main(
        r#"java.util.Properties p = new java.util.Properties(); p.put("k", "old"); p.put("k", "new"); System.out.println(p.get("k"));"#,
    );
    assert_eq!(out, vec!["new"]);
}

#[test]
fn properties_get_or_default() {
    let out = run_main(
        r#"java.util.Properties p = new java.util.Properties(); System.out.println(p.getOrDefault("z", "def"));"#,
    );
    assert_eq!(out, vec!["def"]);
}

#[test]
fn properties_clear_empties() {
    let out = run_main(
        r#"java.util.Properties p = new java.util.Properties(); p.setProperty("a", "1"); p.clear(); System.out.println(p.isEmpty());"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn properties_put_all_copies() {
    let out = run_main(
        r#"java.util.Properties a = new java.util.Properties(); a.setProperty("k", "v"); java.util.Properties b = new java.util.Properties(); b.putAll(a); System.out.println(b.getProperty("k"));"#,
    );
    assert_eq!(out, vec!["v"]);
}

#[test]
fn system_getenv_missing_returns_null() {
    let out = run_main(
        r#"System.out.println(System.getenv("VYBE_TEST_NONEXISTENT_ENV_VAR_XYZ") == null);"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn system_getenv_with_default_pattern() {
    let out = run_main(
        r#"String v = System.getenv("VYBE_TEST_NONEXISTENT_ENV_VAR_XYZ"); System.out.println(v != null ? v : "none");"#,
    );
    assert_eq!(out, vec!["none"]);
}

#[test]
fn system_get_property_line_separator_not_null() {
    let out = run_main(r#"System.out.println(System.getProperty("line.separator") != null);"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn system_get_property_file_separator() {
    let out = run_main(r#"System.out.println(System.getProperty("file.separator").length() > 0);"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn system_get_property_path_separator() {
    let out = run_main(r#"System.out.println(System.getProperty("path.separator").length() > 0);"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn system_get_property_os_name() {
    let out = run_main(r#"System.out.println(System.getProperty("os.name") != null);"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn system_get_property_java_version() {
    let out = run_main(r#"System.out.println(System.getProperty("java.version") != null);"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn system_get_property_user_dir() {
    let out = run_main(r#"System.out.println(System.getProperty("user.dir") != null);"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn system_get_property_user_home() {
    let out = run_main(r#"System.out.println(System.getProperty("user.home") != null);"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn system_get_property_with_default() {
    let out = run_main(r#"System.out.println(System.getProperty("vybe.absent.key", "default"));"#);
    assert_eq!(out, vec!["default"]);
}

#[test]
fn system_set_property_roundtrip() {
    let out = run_main(
        r#"System.setProperty("vybe.test.key", "hello"); System.out.println(System.getProperty("vybe.test.key"));"#,
    );
    assert_eq!(out, vec!["hello"]);
}

#[test]
fn system_clear_property_removes() {
    let out = run_main(
        r#"System.setProperty("vybe.temp.key", "x"); System.clearProperty("vybe.temp.key"); System.out.println(System.getProperty("vybe.temp.key") == null);"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn system_get_properties_not_null() {
    let out = run_main(r#"System.out.println(System.getProperties() != null);"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn system_get_properties_contains_set() {
    let out = run_main(
        r#"System.out.println(System.getProperties().getClass().getName().contains("Properties"));"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn properties_store_defaults() {
    let out = run_main(
        r#"java.util.Properties defaults = new java.util.Properties(); defaults.setProperty("d", "def"); java.util.Properties p = new java.util.Properties(defaults); System.out.println(p.getProperty("d"));"#,
    );
    assert_eq!(out, vec!["def"]);
}

#[test]
fn properties_defaults_fallback() {
    let out = run_main(
        r#"java.util.Properties defaults = new java.util.Properties(); defaults.setProperty("k", "fromDefault"); java.util.Properties p = new java.util.Properties(defaults); System.out.println(p.getProperty("k"));"#,
    );
    assert_eq!(out, vec!["fromDefault"]);
}

#[test]
fn properties_override_default() {
    let out = run_main(
        r#"java.util.Properties defaults = new java.util.Properties(); defaults.setProperty("k", "def"); java.util.Properties p = new java.util.Properties(defaults); p.setProperty("k", "own"); System.out.println(p.getProperty("k"));"#,
    );
    assert_eq!(out, vec!["own"]);
}

#[test]
fn properties_keys_enumeration() {
    let out = run_main(
        r#"java.util.Properties p = new java.util.Properties(); p.setProperty("a", "1"); System.out.println(p.keys().hasMoreElements());"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn properties_elements_enumeration() {
    let out = run_main(
        r#"java.util.Properties p = new java.util.Properties(); p.setProperty("a", "1"); System.out.println(p.elements().hasMoreElements());"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn properties_entry_set_size() {
    let out = run_main(
        r#"java.util.Properties p = new java.util.Properties(); p.setProperty("a", "1"); p.setProperty("b", "2"); System.out.println(p.entrySet().size());"#,
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn properties_clone_independent() {
    let out = run_main(
        r#"java.util.Properties p = new java.util.Properties(); p.setProperty("k", "v"); java.util.Properties c = (java.util.Properties) p.clone(); c.setProperty("k", "changed"); System.out.println(p.getProperty("k"));"#,
    );
    assert_eq!(out, vec!["v"]);
}

#[test]
fn system_getenv_returns_string_or_null() {
    let out = run_main(
        r#"Object v = System.getenv("PATH"); System.out.println(v == null ? true : v.getClass() == String.class);"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn system_get_property_java_vendor() {
    let out = run_main(r#"System.out.println(System.getProperty("java.vendor") != null);"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn system_get_property_file_encoding() {
    let out = run_main(r#"System.out.println(System.getProperty("file.encoding") != null);"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn properties_set_property_null_value() {
    let out = run_main(
        r#"java.util.Properties p = new java.util.Properties(); p.setProperty("n", "null"); System.out.println(p.getProperty("n"));"#,
    );
    assert_eq!(out, vec!["null"]);
}

#[test]
fn properties_replace_existing() {
    let out = run_main(
        r#"java.util.Properties p = new java.util.Properties(); p.setProperty("r", "old"); p.replace("r", "old", "new"); System.out.println(p.getProperty("r"));"#,
    );
    assert_eq!(out, vec!["new"]);
}

#[test]
fn properties_compute_if_absent() {
    let out = run_main(
        r#"java.util.Properties p = new java.util.Properties(); p.computeIfAbsent("c", k -> "computed"); System.out.println(p.getProperty("c"));"#,
    );
    assert_eq!(out, vec!["computed"]);
}

#[test]
fn properties_put_if_absent() {
    let out = run_main(
        r#"java.util.Properties p = new java.util.Properties(); p.putIfAbsent("p", "first"); System.out.println(p.get("p"));"#,
    );
    assert_eq!(out, vec!["first"]);
}

#[test]
fn properties_merge_function() {
    let out = run_main(
        r#"java.util.Properties p = new java.util.Properties(); p.put("m", "a"); p.merge("m", "b", (o, n) -> o + n); System.out.println(p.get("m"));"#,
    );
    assert_eq!(out, vec!["ab"]);
}

#[test]
fn system_set_property_returns_previous() {
    let out = run_main(
        r#"System.setProperty("vybe.prev.test", "a"); String prev = System.setProperty("vybe.prev.test", "b"); System.out.println("a".equals(prev));"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn system_getenv_empty_key_null() {
    let out = run_main(r#"System.out.println(System.getenv("") == null);"#);
    assert_eq!(out, vec!["true"]);
}
