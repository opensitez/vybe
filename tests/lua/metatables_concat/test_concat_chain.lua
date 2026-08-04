-- vybe-test: lua/metatables_concat/test_concat_chain
-- origin: languages/lua/tests/lua/test_metatables_concat.rs

local __w1 = "121"
local __i = 0

local mt={__concat=function(a,b) local av = type(a)=='table' and a.v or a; local bv = type(b)=='table' and b.v or b; return av..bv end}; local t1=setmetatable({v='1'}, mt); local t2=setmetatable({v='2'}, mt); do local __t = tostring(t1..t2..t1); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
