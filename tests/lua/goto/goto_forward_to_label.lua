-- vybe-test: lua/goto/goto_forward_to_label
-- origin: languages/lua/tests/lua/test_goto.rs

local __w1 = "2"
local __i = 0

local s = ""
 goto two
 s = s .. "1"
 ::two::
 s = s .. "2"
 do local __t = tostring(s); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
