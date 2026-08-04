-- vybe-test: lua/lexical_scoping_advanced/lexical_for_loop_capture
-- origin: languages/lua/tests/lua/test_lexical_scoping_advanced.rs

local __w1 = "123"
local __i = 0

local fns = {}
for i = 1, 3 do
  fns[i] = function() return i end
end
do local __t = tostring(fns[1]() .. fns[2]() .. fns[3]()); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
