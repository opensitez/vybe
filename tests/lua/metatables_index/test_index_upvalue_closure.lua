-- vybe-test: lua/metatables_index/test_index_upvalue_closure
-- origin: languages/lua/tests/lua/test_metatables_index.rs

local __w1 = "1 2"
local __i = 0

local count=0; local t=setmetatable({}, {__index=function() count=count+1; return count end}); do local __t = tostring(t.a..' '..t.b); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
