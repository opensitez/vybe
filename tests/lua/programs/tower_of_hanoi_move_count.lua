-- vybe-test: lua/programs/tower_of_hanoi_move_count
-- origin: languages/lua/tests/lua/test_programs.rs

local __w1 = "15"
local __i = 0

local function hanoi(n)
  if n == 1 then return 1 end
  return 2 * hanoi(n - 1) + 1
end
do local __t = tostring(hanoi(4)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
