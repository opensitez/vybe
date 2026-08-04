-- vybe-test: lua/environment/env_metatable_index_fallback
-- origin: languages/lua/tests/lua/test_environment.rs

local __w1 = "1"
local __i = 0

local base = {a = 1}
local env = setmetatable({}, {__index = base})
local _ENV = env
do local __t = tostring(a); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
