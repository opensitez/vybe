-- vybe-test: lua/do_blocks/do_loop_iter_scope
-- origin: languages/lua/tests/lua/test_do_blocks.rs

local __w1 = "1,3"
local __i = 0

local closures = {}
for i = 1, 3 do
  do
    local x = i
    closures[i] = function() return x end
  end
end
do local __t = tostring(closures[1]() .. "," .. closures[3]()); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
