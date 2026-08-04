-- vybe-test: lua/iteration/iterator_state_encapsulated_in_closure
-- origin: languages/lua/tests/lua/test_iteration.rs

local __w1 = "3,4,5,6"
local __i = 0

local function range(from, to)
  local i = from - 1
  return function()
    i = i + 1
    if i <= to then return i end
  end
end
local t = {}
for v in range(3, 6) do t[#t+1] = v end
do local __t = tostring(table.concat(t, ',')); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
