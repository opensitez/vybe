-- vybe-test: lua/string_matching_captures/match_date
-- origin: languages/lua/tests/lua/test_string_matching_captures.rs

local __w1 = "2024,07,11"
local __i = 0

local y,m,d = string.match("2024-07-11", "(%d%d%d%d)-(%d%d)-(%d%d)")
do local __t = tostring(y .. "," .. m .. "," .. d); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
