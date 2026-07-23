use super::helpers::run_lua_one;

#[test]
fn test_debug_getupvalue_snapshot_baseline() {
    assert_eq!(
        run_lua_one(
            r#"local up = 1
local function f()
  return up
end
local name, value = debug.getupvalue(f, 1)
print(name == "up" and value == 1)"#
        ),
        "true"
    );
}

#[test]
fn test_debug_getupvalue_snapshot_simple() {
    assert_eq!(
        run_lua_one(
            r#"local up = 2
local function f()
  return up
end
local name, value = debug.getupvalue(f, 1)
print(name == "up" and value == 2)"#
        ),
        "true"
    );
}

#[test]
fn test_debug_getupvalue_snapshot_trimmed() {
    assert_eq!(
        run_lua_one(
            r#"local up = 3
local function f()
  return up
end
local name, value = debug.getupvalue(f, 1)
print(name == "up" and value == 3)"#
        ),
        "true"
    );
}

#[test]
fn test_debug_getupvalue_snapshot_decimal() {
    assert_eq!(
        run_lua_one(
            r#"local up = 4
local function f()
  return up
end
local name, value = debug.getupvalue(f, 1)
print(name == "up" and value == 4)"#
        ),
        "true"
    );
}

#[test]
fn test_debug_getupvalue_snapshot_hexed() {
    assert_eq!(
        run_lua_one(
            r#"local up = 5
local function f()
  return up
end
local name, value = debug.getupvalue(f, 1)
print(name == "up" and value == 5)"#
        ),
        "true"
    );
}

#[test]
fn test_debug_getupvalue_snapshot_prefixed() {
    assert_eq!(
        run_lua_one(
            r#"local up = 6
local function f()
  return up
end
local name, value = debug.getupvalue(f, 1)
print(name == "up" and value == 6)"#
        ),
        "true"
    );
}

#[test]
fn test_debug_getupvalue_snapshot_negative() {
    assert_eq!(
        run_lua_one(
            r#"local up = 7
local function f()
  return up
end
local name, value = debug.getupvalue(f, 1)
print(name == "up" and value == 7)"#
        ),
        "true"
    );
}

#[test]
fn test_debug_getupvalue_snapshot_rounded() {
    assert_eq!(
        run_lua_one(
            r#"local up = 8
local function f()
  return up
end
local name, value = debug.getupvalue(f, 1)
print(name == "up" and value == 8)"#
        ),
        "true"
    );
}

#[test]
fn test_debug_getupvalue_snapshot_offset() {
    assert_eq!(
        run_lua_one(
            r#"local up = 9
local function f()
  return up
end
local name, value = debug.getupvalue(f, 1)
print(name == "up" and value == 9)"#
        ),
        "true"
    );
}

#[test]
fn test_debug_getupvalue_snapshot_paired() {
    assert_eq!(
        run_lua_one(
            r#"local up = 10
local function f()
  return up
end
local name, value = debug.getupvalue(f, 1)
print(name == "up" and value == 10)"#
        ),
        "true"
    );
}

#[test]
fn test_debug_getupvalue_snapshot_nested() {
    assert_eq!(
        run_lua_one(
            r#"local up = 11
local function f()
  return up
end
local name, value = debug.getupvalue(f, 1)
print(name == "up" and value == 11)"#
        ),
        "true"
    );
}

#[test]
fn test_debug_getupvalue_snapshot_metaflow() {
    assert_eq!(
        run_lua_one(
            r#"local up = 12
local function f()
  return up
end
local name, value = debug.getupvalue(f, 1)
print(name == "up" and value == 12)"#
        ),
        "true"
    );
}

#[test]
fn test_debug_getupvalue_snapshot_guarded() {
    assert_eq!(
        run_lua_one(
            r#"local up = 13
local function f()
  return up
end
local name, value = debug.getupvalue(f, 1)
print(name == "up" and value == 13)"#
        ),
        "true"
    );
}

#[test]
fn test_debug_getupvalue_snapshot_mapped() {
    assert_eq!(
        run_lua_one(
            r#"local up = 14
local function f()
  return up
end
local name, value = debug.getupvalue(f, 1)
print(name == "up" and value == 14)"#
        ),
        "true"
    );
}

#[test]
fn test_debug_getupvalue_snapshot_captured() {
    assert_eq!(
        run_lua_one(
            r#"local up = 15
local function f()
  return up
end
local name, value = debug.getupvalue(f, 1)
print(name == "up" and value == 15)"#
        ),
        "true"
    );
}

#[test]
fn test_debug_getupvalue_snapshot_edge_first() {
    assert_eq!(
        run_lua_one(
            r#"local up = 16
local function f()
  return up
end
local name, value = debug.getupvalue(f, 1)
print(name == "up" and value == 16)"#
        ),
        "true"
    );
}

#[test]
fn test_debug_getupvalue_snapshot_edge_second() {
    assert_eq!(
        run_lua_one(
            r#"local up = 17
local function f()
  return up
end
local name, value = debug.getupvalue(f, 1)
print(name == "up" and value == 17)"#
        ),
        "true"
    );
}

#[test]
fn test_debug_getupvalue_snapshot_edge_last() {
    assert_eq!(
        run_lua_one(
            r#"local up = 18
local function f()
  return up
end
local name, value = debug.getupvalue(f, 1)
print(name == "up" and value == 18)"#
        ),
        "true"
    );
}

#[test]
fn test_debug_getupvalue_snapshot_randomized() {
    assert_eq!(
        run_lua_one(
            r#"local up = 19
local function f()
  return up
end
local name, value = debug.getupvalue(f, 1)
print(name == "up" and value == 19)"#
        ),
        "true"
    );
}

#[test]
fn test_debug_getupvalue_snapshot_unicode_like() {
    assert_eq!(
        run_lua_one(
            r#"local up = 20
local function f()
  return up
end
local name, value = debug.getupvalue(f, 1)
print(name == "up" and value == 20)"#
        ),
        "true"
    );
}
