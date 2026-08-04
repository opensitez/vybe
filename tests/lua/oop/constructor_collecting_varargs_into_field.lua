-- vybe-test: lua/oop/constructor_collecting_varargs_into_field
-- origin: languages/lua/tests/lua/test_oop.rs

local __w1 = "3,20"
local __i = 0

local function make_list(...)
  return {items = {...}}
end
local l = make_list(10, 20, 30)
do local __t = tostring(#l.items .. ',' .. l.items[2]); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
