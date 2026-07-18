use super::helpers::run_lua_one;

#[test]
fn test_string_find_first_baseline() {
    assert_eq!(run_lua_one(r#"print(string.find("alpha0 bravo1 charlie2 delta3 echo4 foxtrot5 golf6 hotel7 india8 juliet9 kilo10 lima11 mike12 november13 oscar14 papa15 quebec16 romeo17 sierra18 tango19 uniform20", "alpha0", 1, true) == 1)"#), "true");
}


#[test]
fn test_string_find_first_simple() {
    assert_eq!(run_lua_one(r#"print(string.find("alpha0 bravo1 charlie2 delta3 echo4 foxtrot5 golf6 hotel7 india8 juliet9 kilo10 lima11 mike12 november13 oscar14 papa15 quebec16 romeo17 sierra18 tango19 uniform20", "bravo1", 1, true) == 8)"#), "true");
}


#[test]
fn test_string_find_first_trimmed() {
    assert_eq!(run_lua_one(r#"print(string.find("alpha0 bravo1 charlie2 delta3 echo4 foxtrot5 golf6 hotel7 india8 juliet9 kilo10 lima11 mike12 november13 oscar14 papa15 quebec16 romeo17 sierra18 tango19 uniform20", "charlie2", 1, true) == 15)"#), "true");
}


#[test]
fn test_string_find_first_decimal() {
    assert_eq!(run_lua_one(r#"print(string.find("alpha0 bravo1 charlie2 delta3 echo4 foxtrot5 golf6 hotel7 india8 juliet9 kilo10 lima11 mike12 november13 oscar14 papa15 quebec16 romeo17 sierra18 tango19 uniform20", "delta3", 1, true) == 24)"#), "true");
}


#[test]
fn test_string_find_first_hexed() {
    assert_eq!(run_lua_one(r#"print(string.find("alpha0 bravo1 charlie2 delta3 echo4 foxtrot5 golf6 hotel7 india8 juliet9 kilo10 lima11 mike12 november13 oscar14 papa15 quebec16 romeo17 sierra18 tango19 uniform20", "echo4", 1, true) == 31)"#), "true");
}


#[test]
fn test_string_find_first_prefixed() {
    assert_eq!(run_lua_one(r#"print(string.find("alpha0 bravo1 charlie2 delta3 echo4 foxtrot5 golf6 hotel7 india8 juliet9 kilo10 lima11 mike12 november13 oscar14 papa15 quebec16 romeo17 sierra18 tango19 uniform20", "foxtrot5", 1, true) == 37)"#), "true");
}


#[test]
fn test_string_find_first_negative() {
    assert_eq!(run_lua_one(r#"print(string.find("alpha0 bravo1 charlie2 delta3 echo4 foxtrot5 golf6 hotel7 india8 juliet9 kilo10 lima11 mike12 november13 oscar14 papa15 quebec16 romeo17 sierra18 tango19 uniform20", "golf6", 1, true) == 46)"#), "true");
}


#[test]
fn test_string_find_first_rounded() {
    assert_eq!(run_lua_one(r#"print(string.find("alpha0 bravo1 charlie2 delta3 echo4 foxtrot5 golf6 hotel7 india8 juliet9 kilo10 lima11 mike12 november13 oscar14 papa15 quebec16 romeo17 sierra18 tango19 uniform20", "hotel7", 1, true) == 52)"#), "true");
}


#[test]
fn test_string_find_first_offset() {
    assert_eq!(run_lua_one(r#"print(string.find("alpha0 bravo1 charlie2 delta3 echo4 foxtrot5 golf6 hotel7 india8 juliet9 kilo10 lima11 mike12 november13 oscar14 papa15 quebec16 romeo17 sierra18 tango19 uniform20", "india8", 1, true) == 59)"#), "true");
}


#[test]
fn test_string_find_first_paired() {
    assert_eq!(run_lua_one(r#"print(string.find("alpha0 bravo1 charlie2 delta3 echo4 foxtrot5 golf6 hotel7 india8 juliet9 kilo10 lima11 mike12 november13 oscar14 papa15 quebec16 romeo17 sierra18 tango19 uniform20", "juliet9", 1, true) == 66)"#), "true");
}


#[test]
fn test_string_find_first_nested() {
    assert_eq!(run_lua_one(r#"print(string.find("alpha0 bravo1 charlie2 delta3 echo4 foxtrot5 golf6 hotel7 india8 juliet9 kilo10 lima11 mike12 november13 oscar14 papa15 quebec16 romeo17 sierra18 tango19 uniform20", "kilo10", 1, true) == 74)"#), "true");
}


#[test]
fn test_string_find_first_metaflow() {
    assert_eq!(run_lua_one(r#"print(string.find("alpha0 bravo1 charlie2 delta3 echo4 foxtrot5 golf6 hotel7 india8 juliet9 kilo10 lima11 mike12 november13 oscar14 papa15 quebec16 romeo17 sierra18 tango19 uniform20", "lima11", 1, true) == 81)"#), "true");
}


#[test]
fn test_string_find_first_guarded() {
    assert_eq!(run_lua_one(r#"print(string.find("alpha0 bravo1 charlie2 delta3 echo4 foxtrot5 golf6 hotel7 india8 juliet9 kilo10 lima11 mike12 november13 oscar14 papa15 quebec16 romeo17 sierra18 tango19 uniform20", "mike12", 1, true) == 88)"#), "true");
}


#[test]
fn test_string_find_first_mapped() {
    assert_eq!(run_lua_one(r#"print(string.find("alpha0 bravo1 charlie2 delta3 echo4 foxtrot5 golf6 hotel7 india8 juliet9 kilo10 lima11 mike12 november13 oscar14 papa15 quebec16 romeo17 sierra18 tango19 uniform20", "november13", 1, true) == 95)"#), "true");
}


#[test]
fn test_string_find_first_captured() {
    assert_eq!(run_lua_one(r#"print(string.find("alpha0 bravo1 charlie2 delta3 echo4 foxtrot5 golf6 hotel7 india8 juliet9 kilo10 lima11 mike12 november13 oscar14 papa15 quebec16 romeo17 sierra18 tango19 uniform20", "oscar14", 1, true) == 106)"#), "true");
}


#[test]
fn test_string_find_first_edge_first() {
    assert_eq!(run_lua_one(r#"print(string.find("alpha0 bravo1 charlie2 delta3 echo4 foxtrot5 golf6 hotel7 india8 juliet9 kilo10 lima11 mike12 november13 oscar14 papa15 quebec16 romeo17 sierra18 tango19 uniform20", "papa15", 1, true) == 114)"#), "true");
}


#[test]
fn test_string_find_first_edge_second() {
    assert_eq!(run_lua_one(r#"print(string.find("alpha0 bravo1 charlie2 delta3 echo4 foxtrot5 golf6 hotel7 india8 juliet9 kilo10 lima11 mike12 november13 oscar14 papa15 quebec16 romeo17 sierra18 tango19 uniform20", "quebec16", 1, true) == 121)"#), "true");
}


#[test]
fn test_string_find_first_edge_last() {
    assert_eq!(run_lua_one(r#"print(string.find("alpha0 bravo1 charlie2 delta3 echo4 foxtrot5 golf6 hotel7 india8 juliet9 kilo10 lima11 mike12 november13 oscar14 papa15 quebec16 romeo17 sierra18 tango19 uniform20", "romeo17", 1, true) == 130)"#), "true");
}


#[test]
fn test_string_find_first_randomized() {
    assert_eq!(run_lua_one(r#"print(string.find("alpha0 bravo1 charlie2 delta3 echo4 foxtrot5 golf6 hotel7 india8 juliet9 kilo10 lima11 mike12 november13 oscar14 papa15 quebec16 romeo17 sierra18 tango19 uniform20", "sierra18", 1, true) == 138)"#), "true");
}


#[test]
fn test_string_find_first_unicode_like() {
    assert_eq!(run_lua_one(r#"print(string.find("alpha0 bravo1 charlie2 delta3 echo4 foxtrot5 golf6 hotel7 india8 juliet9 kilo10 lima11 mike12 november13 oscar14 papa15 quebec16 romeo17 sierra18 tango19 uniform20", "tango19", 1, true) == 147)"#), "true");
}
