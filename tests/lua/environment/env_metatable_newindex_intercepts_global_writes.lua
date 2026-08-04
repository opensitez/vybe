-- vybe-test: lua/environment/env_metatable_newindex_intercepts_global_writes
-- origin: languages/lua/tests/lua/test_environment.rs

local written = {}
local env = setmetatable({print = print}, {
  __newindex = function(t, k, v) written[#written+1] = k; rawset(t, k, v) end
})
local _ENV = env
newvar = 1
print(written[1])
