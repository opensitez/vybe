-- vybe-test: lua/environment_lua52/test_env_lexical_override_no_print
-- origin: languages/lua/tests/lua/test_environment_lua52.rs

local _ENV={a=2}; return a
