-- vybe-test: lua/globals/global_string_and_number_mixed
-- origin: languages/lua/tests/lua/test_globals.rs

local __w1 = "lua5"
local __i = 0

name = "lua"
do local __t = tostring(name .. 5); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
