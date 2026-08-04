-- vybe-test: lua/base_getfenv_setfenv/test_getfenv_setfenv_metaflow
-- origin: languages/lua/tests/lua/test_base_getfenv_setfenv.rs

local __w1 = "true"
local __i = 0

local function f()
  return _G or nil
end
local env = {x = 12}
if type(setfenv) == "function" and type(getfenv) == "function" then
  setfenv(f, env)
  do local __t = tostring(getfenv(f).x == 12); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end
else
  do local __t = tostring(type(f) == "function"); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end
end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
