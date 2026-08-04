-- vybe-test: lua/coroutine_create_resume/test_create_nested_resume
-- origin: languages/lua/tests/lua/test_coroutine_create_resume.rs

local __w1 = "true"
local __i = 0

local inner = coroutine.create(function() return 2 end)
local function outer()
  return coroutine.resume(inner)
end
local ok, r = pcall(outer)
do local __t = tostring(ok and type(r) == "boolean"); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
