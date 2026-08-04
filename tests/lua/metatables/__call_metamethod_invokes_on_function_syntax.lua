-- vybe-test: lua/metatables/__call_metamethod_invokes_on_function_syntax
-- origin: languages/lua/tests/lua/test_metatables.rs

local __w1 = "8"
local __i = 0

local t=setmetatable({}, {__call=function(_,a) return a*2 end})
do local __t = tostring(t(4)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
