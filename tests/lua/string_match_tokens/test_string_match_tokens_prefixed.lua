-- vybe-test: lua/string_match_tokens/test_string_match_tokens_prefixed
-- origin: languages/lua/tests/lua/test_string_match_tokens.rs

local __w1 = "true"
local __i = 0

do local __t = tostring(string.match("a6 b7 c8", "b7") == "b7"); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
