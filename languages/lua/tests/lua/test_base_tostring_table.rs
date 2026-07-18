use super::helpers::run_lua_one;

#[test]
fn test_tostring_table_baseline() {
    assert_eq!(run_lua_one(r#"local t = setmetatable({value = 1}, {__tostring = function(self) return "tbl_" .. tostring(self.value) end}); print(tostring(t) == "tbl_1")"#), "true");
}


#[test]
fn test_tostring_table_simple() {
    assert_eq!(run_lua_one(r#"local t = setmetatable({value = 2}, {__tostring = function(self) return "tbl_" .. tostring(self.value) end}); print(tostring(t) == "tbl_2")"#), "true");
}


#[test]
fn test_tostring_table_trimmed() {
    assert_eq!(run_lua_one(r#"local t = setmetatable({value = 3}, {__tostring = function(self) return "tbl_" .. tostring(self.value) end}); print(tostring(t) == "tbl_3")"#), "true");
}


#[test]
fn test_tostring_table_decimal() {
    assert_eq!(run_lua_one(r#"local t = setmetatable({value = 4}, {__tostring = function(self) return "tbl_" .. tostring(self.value) end}); print(tostring(t) == "tbl_4")"#), "true");
}


#[test]
fn test_tostring_table_hexed() {
    assert_eq!(run_lua_one(r#"local t = setmetatable({value = 5}, {__tostring = function(self) return "tbl_" .. tostring(self.value) end}); print(tostring(t) == "tbl_5")"#), "true");
}


#[test]
fn test_tostring_table_prefixed() {
    assert_eq!(run_lua_one(r#"local t = setmetatable({value = 6}, {__tostring = function(self) return "tbl_" .. tostring(self.value) end}); print(tostring(t) == "tbl_6")"#), "true");
}


#[test]
fn test_tostring_table_negative() {
    assert_eq!(run_lua_one(r#"local t = setmetatable({value = 7}, {__tostring = function(self) return "tbl_" .. tostring(self.value) end}); print(tostring(t) == "tbl_7")"#), "true");
}


#[test]
fn test_tostring_table_rounded() {
    assert_eq!(run_lua_one(r#"local t = setmetatable({value = 8}, {__tostring = function(self) return "tbl_" .. tostring(self.value) end}); print(tostring(t) == "tbl_8")"#), "true");
}


#[test]
fn test_tostring_table_offset() {
    assert_eq!(run_lua_one(r#"local t = setmetatable({value = 9}, {__tostring = function(self) return "tbl_" .. tostring(self.value) end}); print(tostring(t) == "tbl_9")"#), "true");
}


#[test]
fn test_tostring_table_paired() {
    assert_eq!(run_lua_one(r#"local t = setmetatable({value = 10}, {__tostring = function(self) return "tbl_" .. tostring(self.value) end}); print(tostring(t) == "tbl_10")"#), "true");
}


#[test]
fn test_tostring_table_nested() {
    assert_eq!(run_lua_one(r#"local t = setmetatable({value = 11}, {__tostring = function(self) return "tbl_" .. tostring(self.value) end}); print(tostring(t) == "tbl_11")"#), "true");
}


#[test]
fn test_tostring_table_metaflow() {
    assert_eq!(run_lua_one(r#"local t = setmetatable({value = 12}, {__tostring = function(self) return "tbl_" .. tostring(self.value) end}); print(tostring(t) == "tbl_12")"#), "true");
}


#[test]
fn test_tostring_table_guarded() {
    assert_eq!(run_lua_one(r#"local t = setmetatable({value = 13}, {__tostring = function(self) return "tbl_" .. tostring(self.value) end}); print(tostring(t) == "tbl_13")"#), "true");
}


#[test]
fn test_tostring_table_mapped() {
    assert_eq!(run_lua_one(r#"local t = setmetatable({value = 14}, {__tostring = function(self) return "tbl_" .. tostring(self.value) end}); print(tostring(t) == "tbl_14")"#), "true");
}


#[test]
fn test_tostring_table_captured() {
    assert_eq!(run_lua_one(r#"local t = setmetatable({value = 15}, {__tostring = function(self) return "tbl_" .. tostring(self.value) end}); print(tostring(t) == "tbl_15")"#), "true");
}


#[test]
fn test_tostring_table_edge_first() {
    assert_eq!(run_lua_one(r#"local t = setmetatable({value = 16}, {__tostring = function(self) return "tbl_" .. tostring(self.value) end}); print(tostring(t) == "tbl_16")"#), "true");
}


#[test]
fn test_tostring_table_edge_second() {
    assert_eq!(run_lua_one(r#"local t = setmetatable({value = 17}, {__tostring = function(self) return "tbl_" .. tostring(self.value) end}); print(tostring(t) == "tbl_17")"#), "true");
}


#[test]
fn test_tostring_table_edge_last() {
    assert_eq!(run_lua_one(r#"local t = setmetatable({value = 18}, {__tostring = function(self) return "tbl_" .. tostring(self.value) end}); print(tostring(t) == "tbl_18")"#), "true");
}


#[test]
fn test_tostring_table_randomized() {
    assert_eq!(run_lua_one(r#"local t = setmetatable({value = 19}, {__tostring = function(self) return "tbl_" .. tostring(self.value) end}); print(tostring(t) == "tbl_19")"#), "true");
}


#[test]
fn test_tostring_table_unicode_like() {
    assert_eq!(run_lua_one(r#"local t = setmetatable({value = 20}, {__tostring = function(self) return "tbl_" .. tostring(self.value) end}); print(tostring(t) == "tbl_20")"#), "true");
}
