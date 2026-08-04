-- vybe-test: lua/errors/xpcall_yield_inside_function
-- origin: languages/lua/tests/lua/test_errors.rs

local __w1 = "yielding true"
local __i = 0

local co = coroutine.create(function()
  local ok, res = xpcall(function() return coroutine.yield("yielding") end, function(e) return e end)
  return ok
end)
local _, val = coroutine.resume(co)
local _, val2 = coroutine.resume(co)
do local __t = tostring(val .. " " .. tostring(val2)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
