lua_print! {
    test_pattern_balanced_parentheses => {
        "local s = 'a (b (c) d) e'
local m = string.match(s, '%b()')
print(m)",
        "(b (c) d)"
    },
    test_pattern_frontier => {
        "local s = 'the quick brown fox'
local m = string.match(s, '%f[%w]quick%f[%W]')
print(m)",
        "quick"
    },
    test_pattern_frontier_at_start => {
        "local s = 'the quick'
local m = string.match(s, '%f[%w]the%f[%W]')
print(m)",
        "the"
    },
    test_pattern_captures_in_gsub => {
        "local s = 'hello world'
local r = string.gsub(s, '(%w+) (%w+)', '%2 %1')
print(r)",
        "world hello"
    },
    test_pattern_function_in_gsub => {
        "local s = 'a b c'
local r = string.gsub(s, '%w+', function(w) return string.upper(w) end)
print(r)",
        "A B C"
    },
    test_pattern_table_in_gsub => {
        "local t = {a = 'alpha', b = 'beta'}
local r = string.gsub('a and b', '%w+', t)
print(r)",
        "alpha and beta"
    },
    test_pattern_multiple_captures_gmatch => {
        "local s = 'key1=value1 key2=value2'
local res = ''
for k, v in string.gmatch(s, '(%w+)=(%w+)') do
    res = res .. k .. ':' .. v .. ' '
end
print(res)",
        "key1:value1 key2:value2 "
    }
}
