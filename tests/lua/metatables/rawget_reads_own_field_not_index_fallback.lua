-- vybe-test: lua/metatables/rawget_reads_own_field_not_index_fallback
-- origin: languages/lua/tests/lua/test_metatables.rs

local __w1 = "1"
local __i = 0

local t = setmetatable({x = 1}, {__index = {x = 9}})
do local __t = tostring(rawget(t, "x")); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
