-- vybe-test: lua/env_lexical_binding/env_table_update
-- origin: languages/lua/tests/lua/test_env_lexical_binding.rs

local env = {print=print}
do
  local _ENV = env
  myGlobal = 99
end
print(env.myGlobal)
