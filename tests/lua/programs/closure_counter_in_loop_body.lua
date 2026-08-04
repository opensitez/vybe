-- vybe-test: lua/programs/closure_counter_in_loop_body
-- origin: languages/lua/tests/lua/test_programs.rs

local __w1 = "2"
local __i = 0

local fns = {}
for i = 1, 3 do
  fns[i] = function() return i end
end
do local __t = tostring(fns[2]()); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
