-- vybe-test: lua/operators_concat/test_concat_right_associative
-- origin: languages/lua/tests/lua/test_operators_concat.rs

local __w1 = "string"
local __i = 0

local t={}; setmetatable(t, {__concat=function(a,b) return tostring(a)..tostring(b) end}); do local __t = tostring(type(t .. 'a' .. 'b')); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
