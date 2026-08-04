-- vybe-test: lua/metatables/__add_metamethod_on_tables
-- origin: languages/lua/tests/lua/test_metatables.rs

local __w1 = "7"
local __i = 0

local mt={__add=function(a,b) return {v=a.v+b.v} end}
local a=setmetatable({v=2},mt)
local b=setmetatable({v=5},mt)
do local __t = tostring((a+b).v); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
