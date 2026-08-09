use super::helpers::run_python;

// configparser — ConfigParser, read_string, get, set, add_section, has_section, options, items, write, fallback, RawConfigParser, ExtendedInterpolation, getint, getfloat, getboolean

#[test]
fn test_configparser_read_string_and_get() {
    let out = run_python(
        r#"
import configparser
ini = """
[DEFAULT]
server = localhost

[database]
port = 5432
name = mydb
"""
config = configparser.ConfigParser()
config.read_string(ini)
print(config.get("database", "server"))
print(config.get("database", "name"))
"#,
    );
    assert_eq!(out, vec!["localhost", "mydb"]);
}

#[test]
fn test_configparser_typed_getters_int_float_boolean() {
    let out = run_python(
        r#"
import configparser
ini = """
[settings]
port = 8080
ratio = 0.75
debug = yes
enabled = true
offline = 0
"""
config = configparser.ConfigParser()
config.read_string(ini)
print(config.getint("settings", "port"))
print(config.getfloat("settings", "ratio"))
print(config.getboolean("settings", "debug"))
print(config.getboolean("settings", "enabled"))
print(config.getboolean("settings", "offline"))
"#,
    );
    assert_eq!(out, vec!["8080", "0.75", "True", "True", "False"]);
}

#[test]
fn test_configparser_fallback_values() {
    let out = run_python(
        r#"
import configparser
config = configparser.ConfigParser()
config.read_string("[section]")
print(config.get("section", "missing_key", fallback="default_val"))
print(config.getint("section", "missing_int", fallback=42))
"#,
    );
    assert_eq!(out, vec!["default_val", "42"]);
}

#[test]
fn test_configparser_add_section_and_set() {
    let out = run_python(
        r#"
import configparser
config = configparser.ConfigParser()
config.add_section("user")
config.set("user", "name", "Alice")
print(config.has_section("user"))
print(config.get("user", "name"))
"#,
    );
    assert_eq!(out, vec!["True", "Alice"]);
}

#[test]
fn test_configparser_options_and_items() {
    let out = run_python(
        r#"
import configparser
ini = """
[web]
host = 127.0.0.1
port = 80
"""
config = configparser.ConfigParser()
config.read_string(ini)
print(config.options("web"))
print(dict(config.items("web")))
"#,
    );
    assert_eq!(
        out,
        vec!["['host', 'port']", "{'host': '127.0.0.1', 'port': '80'}"]
    );
}

#[test]
fn test_configparser_write_to_string_stream() {
    let out = run_python(
        r#"
import configparser, io
config = configparser.ConfigParser()
config["app"] = {"title": "My Application", "version": "1.0"}
buf = io.StringIO()
config.write(buf)
out = buf.getvalue()
print("[app]" in out)
print("title = My Application" in out)
"#,
    );
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_configparser_basic_interpolation() {
    let out = run_python(
        r#"
import configparser
ini = """
[paths]
base = /var/www
html = %(base)s/html
cgi = %(base)s/cgi-bin
"""
config = configparser.ConfigParser()
config.read_string(ini)
print(config.get("paths", "html"))
print(config.get("paths", "cgi"))
"#,
    );
    assert_eq!(out, vec!["/var/www/html", "/var/www/cgi-bin"]);
}

#[test]
fn test_configparser_extended_interpolation() {
    let out = run_python(
        r#"
import configparser
ini = """
[common]
home = /home/user

[user_paths]
downloads = ${common:home}/Downloads
"""
config = configparser.ConfigParser(interpolation=configparser.ExtendedInterpolation())
config.read_string(ini)
print(config.get("user_paths", "downloads"))
"#,
    );
    assert_eq!(out, vec!["/home/user/Downloads"]);
}

#[test]
fn test_configparser_raw_config_parser_no_interpolation() {
    let out = run_python(
        r#"
import configparser
ini = """
[sec]
pattern = %(not_interpolated)s
"""
config = configparser.RawConfigParser()
config.read_string(ini)
print(config.get("sec", "pattern"))
"#,
    );
    assert_eq!(out, vec!["%(not_interpolated)s"]);
}

#[test]
fn test_configparser_duplicate_section_error() {
    let out = run_python(
        r#"
import configparser
config = configparser.ConfigParser()
config.add_section("sec")
try:
    config.add_section("sec")
except configparser.DuplicateSectionError:
    print("DuplicateSectionError")
"#,
    );
    assert_eq!(out, vec!["DuplicateSectionError"]);
}

#[test]
fn test_configparser_no_section_error() {
    let out = run_python(
        r#"
import configparser
config = configparser.ConfigParser()
try:
    config.get("non_existent", "key")
except configparser.NoSectionError:
    print("NoSectionError")
"#,
    );
    assert_eq!(out, vec!["NoSectionError"]);
}

#[test]
fn test_configparser_no_option_error() {
    let out = run_python(
        r#"
import configparser
config = configparser.ConfigParser()
config.add_section("sec")
try:
    config.get("sec", "missing")
except configparser.NoOptionError:
    print("NoOptionError")
"#,
    );
    assert_eq!(out, vec!["NoOptionError"]);
}

#[test]
fn test_configparser_dictionary_mapping_access() {
    let out = run_python(
        r#"
import configparser
config = configparser.ConfigParser()
config["meta"] = {"author": "John", "license": "MIT"}
print(config["meta"]["author"])
print(config["meta"]["license"])
"#,
    );
    assert_eq!(out, vec!["John", "MIT"]);
}

#[test]
fn test_configparser_remove_option_and_section() {
    let out = run_python(
        r#"
import configparser
config = configparser.ConfigParser()
config.read_string("[s]\na = 1\nb = 2")
config.remove_option("s", "a")
print(config.has_option("s", "a"))
config.remove_section("s")
print(config.has_section("s"))
"#,
    );
    assert_eq!(out, vec!["False", "False"]);
}

#[test]
fn test_configparser_case_insensitive_option_keys() {
    let out = run_python(
        r#"
import configparser
config = configparser.ConfigParser()
config.read_string("[sec]\nKEYNAME = val")
print(config.get("sec", "keyname"))
"#,
    );
    assert_eq!(out, vec!["val"]);
}

#[test]
fn test_configparser_custom_optionxform() {
    let out = run_python(
        r#"
import configparser
config = configparser.ConfigParser()
config.optionxform = str  # Preserve case
config.read_string("[sec]\nKeyName = val")
print(config.has_option("sec", "KeyName"))
print(config.has_option("sec", "keyname"))
"#,
    );
    assert_eq!(out, vec!["True", "False"]);
}

#[test]
fn test_configparser_default_section_inheritance() {
    let out = run_python(
        r#"
import configparser
config = configparser.ConfigParser(default_section="DEFAULT")
config["DEFAULT"]["global_var"] = "123"
config.add_section("sec1")
print(config.get("sec1", "global_var"))
"#,
    );
    assert_eq!(out, vec!["123"]);
}

#[test]
fn test_configparser_read_dict() {
    let out = run_python(
        r#"
import configparser
d = {"sec1": {"k1": "v1"}, "sec2": {"k2": "v2"}}
config = configparser.ConfigParser()
config.read_dict(d)
print(config.get("sec1", "k1"))
print(config.get("sec2", "k2"))
"#,
    );
    assert_eq!(out, vec!["v1", "v2"]);
}

#[test]
fn test_configparser_multiline_values() {
    let out = run_python(
        r#"
import configparser
ini = """
[notes]
text = line1
    line2
    line3
"""
config = configparser.ConfigParser()
config.read_string(ini)
val = config.get("notes", "text")
print("line1\nline2\nline3" in val)
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_configparser_allow_no_value_options() {
    let out = run_python(
        r#"
import configparser
ini = """
[flags]
enable_feature
"""
config = configparser.ConfigParser(allow_no_value=True)
config.read_string(ini)
print(config.get("flags", "enable_feature") is None)
"#,
    );
    assert_eq!(out, vec!["True"]);
}
