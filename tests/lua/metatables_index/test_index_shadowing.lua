-- vybe-test: lua/metatables_index/test_index_shadowing
-- origin: languages/lua/tests/lua/test_metatables_index.rs

local __w1 = "1 3"
local __i = 0

local t=setmetatable({a=1}, {__index={a=2, b=3}}); do local __t = tostring(t.a..' '..t.b); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
