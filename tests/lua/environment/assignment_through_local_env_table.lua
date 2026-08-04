-- vybe-test: lua/environment/assignment_through_local_env_table
-- origin: languages/lua/tests/lua/test_environment.rs

local t = {n = 1, print = print}
local _ENV = t
t.n = 2
print(n)
