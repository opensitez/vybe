-- vybe-test: lua/scoping/scoping_closure_created_inside_loop_captures_correct_variable_instance
-- origin: languages/lua/tests/lua/test_scoping.rs

local __w1 = "1 2"
local __i = 0

local funcs = {}
for i = 1, 2 do
  local val = i
  funcs[i] = function() return val end
end
do local __t = tostring(funcs[1]() .. " " .. funcs[2]()); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
