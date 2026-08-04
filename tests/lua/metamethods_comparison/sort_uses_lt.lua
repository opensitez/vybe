-- vybe-test: lua/metamethods_comparison/sort_uses_lt
-- origin: languages/lua/tests/lua/test_metamethods_comparison.rs

local __w1 = "1,2,3"
local __i = 0

local mt = {__lt = function(a, b) return a.v < b.v end}
local function W(n) return setmetatable({v=n}, mt) end
local t = {W(3), W(1), W(2)}
table.sort(t)
do local __t = tostring(t[1].v .. "," .. t[2].v .. "," .. t[3].v); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
