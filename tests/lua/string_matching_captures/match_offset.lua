-- vybe-test: lua/string_matching_captures/match_offset
-- origin: languages/lua/tests/lua/test_string_matching_captures.rs

local __w1 = "bb"
local __i = 0

do local __t = tostring(string.match("aabba", "b+", 3)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
