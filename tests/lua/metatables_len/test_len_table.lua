-- vybe-test: lua/metatables_len/test_len_table
-- origin: languages/lua/tests/lua/test_metatables_len.rs

local __w1 = "42"
local __i = 0

local mt={__len=function(t) return 42 end}; local t=setmetatable({1,2,3}, mt); do local __t = tostring(#t); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
