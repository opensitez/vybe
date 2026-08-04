-- vybe-test: lua/coroutines_advanced/test_coroutine_iterator_pattern
-- origin: languages/lua/tests/lua/test_coroutines_advanced.rs

local __w1 = "6"
local __i = 0

local function permgen(a, n)
    n = n or #a
    if n <= 1 then
        coroutine.yield(a)
    else
        for i = 1, n do
            a[n], a[i] = a[i], a[n]
            permgen(a, n - 1)
            a[n], a[i] = a[i], a[n]
        end
    end
end
local function permutations(a)
    local co = coroutine.create(function() permgen(a) end)
    return function()
        local code, res = coroutine.resume(co)
        return res
    end
end
local count = 0
for p in permutations({1, 2, 3}) do
    count = count + 1
end
do local __t = tostring(count); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
