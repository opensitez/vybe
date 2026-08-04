-- vybe-test: lua/metatables/__len_metamethod_overrides_length_operator
-- origin: languages/lua/tests/lua/test_metatables.rs

local __w1 = "5"
local __i = 0

local t=setmetatable({},{__len=function() return 5 end})
do local __t = tostring(#t); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
