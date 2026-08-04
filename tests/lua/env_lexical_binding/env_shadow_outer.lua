-- vybe-test: lua/env_lexical_binding/env_shadow_outer
-- origin: languages/lua/tests/lua/test_env_lexical_binding.rs

local outer_env = _ENV
local _ENV = {print=print, outer=outer_env}
outer.print("ok")
