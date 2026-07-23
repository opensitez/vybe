use super::helpers::run_lua_one;

#[test]
fn test_table_metamap_access_baseline() {
    assert_eq!(
        run_lua_one(
            r#"local mt = {__index = function(_, k) return k end}
local t = setmetatable({}, mt)
print(t["k1"] == "k1")"#
        ),
        "true"
    );
}

#[test]
fn test_table_metamap_access_simple() {
    assert_eq!(
        run_lua_one(
            r#"local mt = {__index = function(_, k) return k end}
local t = setmetatable({}, mt)
print(t["k2"] == "k2")"#
        ),
        "true"
    );
}

#[test]
fn test_table_metamap_access_trimmed() {
    assert_eq!(
        run_lua_one(
            r#"local mt = {__index = function(_, k) return k end}
local t = setmetatable({}, mt)
print(t["k3"] == "k3")"#
        ),
        "true"
    );
}

#[test]
fn test_table_metamap_access_decimal() {
    assert_eq!(
        run_lua_one(
            r#"local mt = {__index = function(_, k) return k end}
local t = setmetatable({}, mt)
print(t["k4"] == "k4")"#
        ),
        "true"
    );
}

#[test]
fn test_table_metamap_access_hexed() {
    assert_eq!(
        run_lua_one(
            r#"local mt = {__index = function(_, k) return k end}
local t = setmetatable({}, mt)
print(t["k5"] == "k5")"#
        ),
        "true"
    );
}

#[test]
fn test_table_metamap_access_prefixed() {
    assert_eq!(
        run_lua_one(
            r#"local mt = {__index = function(_, k) return k end}
local t = setmetatable({}, mt)
print(t["k6"] == "k6")"#
        ),
        "true"
    );
}

#[test]
fn test_table_metamap_access_negative() {
    assert_eq!(
        run_lua_one(
            r#"local mt = {__index = function(_, k) return k end}
local t = setmetatable({}, mt)
print(t["k7"] == "k7")"#
        ),
        "true"
    );
}

#[test]
fn test_table_metamap_access_rounded() {
    assert_eq!(
        run_lua_one(
            r#"local mt = {__index = function(_, k) return k end}
local t = setmetatable({}, mt)
print(t["k8"] == "k8")"#
        ),
        "true"
    );
}

#[test]
fn test_table_metamap_access_offset() {
    assert_eq!(
        run_lua_one(
            r#"local mt = {__index = function(_, k) return k end}
local t = setmetatable({}, mt)
print(t["k9"] == "k9")"#
        ),
        "true"
    );
}

#[test]
fn test_table_metamap_access_paired() {
    assert_eq!(
        run_lua_one(
            r#"local mt = {__index = function(_, k) return k end}
local t = setmetatable({}, mt)
print(t["k10"] == "k10")"#
        ),
        "true"
    );
}

#[test]
fn test_table_metamap_access_nested() {
    assert_eq!(
        run_lua_one(
            r#"local mt = {__index = function(_, k) return k end}
local t = setmetatable({}, mt)
print(t["k11"] == "k11")"#
        ),
        "true"
    );
}

#[test]
fn test_table_metamap_access_metaflow() {
    assert_eq!(
        run_lua_one(
            r#"local mt = {__index = function(_, k) return k end}
local t = setmetatable({}, mt)
print(t["k12"] == "k12")"#
        ),
        "true"
    );
}

#[test]
fn test_table_metamap_access_guarded() {
    assert_eq!(
        run_lua_one(
            r#"local mt = {__index = function(_, k) return k end}
local t = setmetatable({}, mt)
print(t["k13"] == "k13")"#
        ),
        "true"
    );
}

#[test]
fn test_table_metamap_access_mapped() {
    assert_eq!(
        run_lua_one(
            r#"local mt = {__index = function(_, k) return k end}
local t = setmetatable({}, mt)
print(t["k14"] == "k14")"#
        ),
        "true"
    );
}

#[test]
fn test_table_metamap_access_captured() {
    assert_eq!(
        run_lua_one(
            r#"local mt = {__index = function(_, k) return k end}
local t = setmetatable({}, mt)
print(t["k15"] == "k15")"#
        ),
        "true"
    );
}

#[test]
fn test_table_metamap_access_edge_first() {
    assert_eq!(
        run_lua_one(
            r#"local mt = {__index = function(_, k) return k end}
local t = setmetatable({}, mt)
print(t["k16"] == "k16")"#
        ),
        "true"
    );
}

#[test]
fn test_table_metamap_access_edge_second() {
    assert_eq!(
        run_lua_one(
            r#"local mt = {__index = function(_, k) return k end}
local t = setmetatable({}, mt)
print(t["k17"] == "k17")"#
        ),
        "true"
    );
}

#[test]
fn test_table_metamap_access_edge_last() {
    assert_eq!(
        run_lua_one(
            r#"local mt = {__index = function(_, k) return k end}
local t = setmetatable({}, mt)
print(t["k18"] == "k18")"#
        ),
        "true"
    );
}

#[test]
fn test_table_metamap_access_randomized() {
    assert_eq!(
        run_lua_one(
            r#"local mt = {__index = function(_, k) return k end}
local t = setmetatable({}, mt)
print(t["k19"] == "k19")"#
        ),
        "true"
    );
}

#[test]
fn test_table_metamap_access_unicode_like() {
    assert_eq!(
        run_lua_one(
            r#"local mt = {__index = function(_, k) return k end}
local t = setmetatable({}, mt)
print(t["k20"] == "k20")"#
        ),
        "true"
    );
}
