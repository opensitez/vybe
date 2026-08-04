-- vybe-test: lua/metatables_index/test_index_metamethod_type_number
-- origin: languages/lua/tests/lua/test_metatables_index.rs

local __w1 = "15"
local __i = 0

debug.setmetatable(0, {__index=function(n, k) return n+k end}); do local __t = tostring((10)[5]); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
