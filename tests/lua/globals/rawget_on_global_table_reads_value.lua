-- vybe-test: lua/globals/rawget_on_global_table_reads_value
-- origin: languages/lua/tests/lua/test_globals.rs

local __w1 = "7"
local __i = 0

xyzzy = 7
do local __t = tostring(rawget(_G, "xyzzy")); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
