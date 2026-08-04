-- vybe-test: lua/coroutine_wrappers/test_wrap_error_propagation
-- origin: languages/lua/tests/lua/test_coroutine_wrappers.rs

local __w1 = "false true"
local __i = 0

local f = coroutine.wrap(function()
    error('wrap error')
end)
local ok, err = pcall(f)
do local __t = tostring(tostring(ok) .. ' ' .. tostring(string.find(err, 'wrap error') ~= nil)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
