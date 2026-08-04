-- vybe-test: lua/tables_custom_iterators/test_iterator_fibonacci
-- origin: languages/lua/tests/lua/test_tables_custom_iterators.rs

local __w1 = "0 1 1 2 3 5 8 "
local __i = 0

local function fib(max)
    local a, b = 0, 1
    return function()
        if a > max then return nil end
        local curr = a
        a, b = b, a + b
        return curr
    end
end
local s = ''
for v in fib(10) do
    s = s .. v .. ' '
end
do local __t = tostring(s); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
