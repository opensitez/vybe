use super::helpers::run_lua_one;

#[test]
fn test_string_gmatch_empty_matches() {
    assert_eq!(
        run_lua_one("local s='' for w in string.gmatch('a', '.*') do s=s..w..',' end print(s)"),
        "a, "
    );
}

#[test]
fn test_string_gmatch_multiple_captures() {
    assert_eq!(
        run_lua_one("local s='' for k,v in string.gmatch('a=1 b=2', '(%w+)=(%d+)') do s=s..k..v end print(s)"),
        "a1b2"
    );
}

#[test]
fn test_string_gsub_with_table() {
    assert_eq!(
        run_lua_one("local t={a='A', b='B'} print((string.gsub('a b c', '%w', t)))"),
        "A B c"
    );
}

#[test]
fn test_string_gsub_with_function_returning_nil() {
    assert_eq!(
        run_lua_one("print((string.gsub('hello world', '%w+', function(w) if w=='world' then return 'LUA' end end)))"),
        "hello LUA"
    );
}

#[test]
fn test_string_gsub_with_function_multiple_args() {
    assert_eq!(
        run_lua_one("print((string.gsub('x=10, y=20', '(%w+)=(%d+)', function(k,v) return k..tonumber(v)*2 end)))"),
        "x20, y40"
    );
}

#[test]
fn test_string_gsub_count_limit() {
    assert_eq!(
        run_lua_one("local res, cnt = string.gsub('a a a a a', 'a', 'b', 3) print(res .. ' ' .. cnt)"),
        "b b b a a 3"
    );
}

#[test]
fn test_string_find_plain_mode() {
    assert_eq!(
        run_lua_one("local s, e = string.find('a-b-c', '-', 1, true) print(s .. ' ' .. e)"),
        "2 2"
    );
}

#[test]
fn test_string_find_position_capture() {
    assert_eq!(
        run_lua_one("local s, e, p1, p2 = string.find('hello', '()ll()') print(p1 .. ' ' .. p2)"),
        "3 5"
    );
}

#[test]
fn test_string_match_frontier_pattern() {
    assert_eq!(
        run_lua_one("print((string.match('the quick brown', '%f[%w]quick%f[%W]')))"),
        "quick"
    );
}

#[test]
fn test_string_match_frontier_pattern_at_string_ends() {
    assert_eq!(
        run_lua_one("print(string.match('test', '%f[%w]test%f[%W]') ~= nil)"),
        "true"
    );
}

#[test]
fn test_string_match_balanced_captures() {
    assert_eq!(
        run_lua_one("print((string.match('abc(def(ghi)jkl)mno', '%b()')))"),
        "(def(ghi)jkl)"
    );
}

#[test]
fn test_string_match_balanced_custom_characters() {
    assert_eq!(
        run_lua_one("print((string.match('x <y <z> w> v', '%b<>')))"),
        "<y <z> w>"
    );
}

#[test]
fn test_string_character_classes_digits() {
    assert_eq!(
        run_lua_one("print((string.gsub('a1b2c3D4', '%D', '')))"),
        "1234"
    );
}

#[test]
fn test_string_character_classes_hex() {
    assert_eq!(
        run_lua_one("print((string.gsub('g1hA2zF', '%X', '')))"),
        "1A2F"
    );
}

#[test]
fn test_string_character_classes_punctuation() {
    assert_eq!(
        run_lua_one("print((string.gsub('a!b,c.d e', '%p', 'X')))"),
        "aXbXcXd e"
    );
}

#[test]
fn test_string_match_optional_character() {
    assert_eq!(
        run_lua_one("print((string.match('color or colour', 'colou?r')))"),
        "color"
    );
}

#[test]
fn test_string_match_optional_capture() {
    assert_eq!(
        run_lua_one("local a, b = string.match('http://example.com', '^(https?)://(.+)$') print(a .. ' ' .. b)"),
        "http example.com"
    );
}

#[test]
fn test_string_match_magic_characters_escaped() {
    assert_eq!(
        run_lua_one("print((string.match('a-b*c+d?e^f$g(h)i[j]k%l.m', '%-%*%+%?%^%$%(%)%[%]%%%.')))"),
        "-*+?^$()[]%."
    );
}

#[test]
fn test_string_match_character_set_complement() {
    assert_eq!(
        run_lua_one("print((string.gsub('a1b2c3', '[^a-z]', '')))"),
        "abc"
    );
}

#[test]
fn test_string_match_character_set_with_escapes() {
    assert_eq!(
        run_lua_one("print((string.match('a]b-c', '[%a%]%-]+')))"),
        "a]b-c"
    );
}

#[test]
fn test_string_gsub_reference_captures() {
    assert_eq!(
        run_lua_one("print((string.gsub('hello', '(.)(.)', '%2%1')))"),
        "ehll o"
    );
}

#[test]
fn test_string_gsub_reference_entire_match() {
    assert_eq!(
        run_lua_one("print((string.gsub('abc', '%a', '<%0>')))"),
        "<a><b><c>"
    );
}
