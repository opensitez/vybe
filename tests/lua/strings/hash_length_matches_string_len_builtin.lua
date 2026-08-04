-- vybe-test: lua/strings/hash_length_matches_string_len_builtin
-- origin: languages/lua/tests/lua/test_strings.rs

local __w1 = "true"
local __i = 0

do local __t = tostring(#"hello" == string.len("hello")); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
