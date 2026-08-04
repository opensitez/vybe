-- vybe-test: lua/closures_complex/test_closures_captured_loop_variables
-- origin: languages/lua/tests/lua/test_closures_complex.rs

local __w1 = "1 2 3"
local __i = 0

local funcs = {}
for i = 1, 3 do
    local v = i
    funcs[i] = function() return v end
end
do local __t = tostring(funcs[1]() .. ' ' .. funcs[2]() .. ' ' .. funcs[3]()); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
