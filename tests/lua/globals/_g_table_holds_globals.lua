-- vybe-test: lua/globals/_g_table_holds_globals
-- origin: languages/lua/tests/lua/test_globals.rs

local __w1 = "42"
local __i = 0

_G.answer = 42
do local __t = tostring(_G.answer); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
