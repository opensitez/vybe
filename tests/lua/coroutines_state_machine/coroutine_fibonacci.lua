-- vybe-test: lua/coroutines_state_machine/coroutine_fibonacci
-- origin: languages/lua/tests/lua/test_coroutines_state_machine.rs

local __w1 = "0,1,1,2,3,5,8"
local __i = 0

local function fib_gen()
  local a, b = 0, 1
  while true do
    coroutine.yield(a)
    a, b = b, a + b
  end
end
local co = coroutine.create(fib_gen)
local nums = {}
for _ = 1, 7 do
  local _, v = coroutine.resume(co)
  nums[#nums+1] = v
end
do local __t = tostring(table.concat(nums, ",")); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
