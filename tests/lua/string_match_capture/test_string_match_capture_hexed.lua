-- vybe-test: lua/string_match_capture/test_string_match_capture_hexed
-- origin: languages/lua/tests/lua/test_string_match_capture.rs

local __w1 = "true"
local __i = 0

local tag, num = string.match("id5:value10", "(id%d+):(value%d+)")
do local __t = tostring(tag == "id5" and num == "value10"); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
