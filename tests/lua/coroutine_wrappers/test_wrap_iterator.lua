-- vybe-test: lua/coroutine_wrappers/test_wrap_iterator
-- origin: languages/lua/tests/lua/test_coroutine_wrappers.rs

local __w1 = "123"
local __i = 0

local function traverse(t)
    return coroutine.wrap(function()
        for i, v in ipairs(t) do
            coroutine.yield(v)
        end
    end)
end
local s = ''
for v in traverse({1, 2, 3}) do
    s = s .. v
end
do local __t = tostring(s); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
