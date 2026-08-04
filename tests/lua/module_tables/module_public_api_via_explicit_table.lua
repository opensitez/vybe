-- vybe-test: lua/module_tables/module_public_api_via_explicit_table
-- origin: languages/lua/tests/lua/test_module_tables.rs

local __w1 = "pub:secret"
local __i = 0

local function create_module()
  local private = 'secret'
  return {
    public = function() return 'pub:' .. private end
  }
end
local mod = create_module()
do local __t = tostring(mod.public()); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
