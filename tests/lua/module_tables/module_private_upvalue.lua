-- vybe-test: lua/module_tables/module_private_upvalue
-- origin: languages/lua/tests/lua/test_module_tables.rs

local __w1 = "3"
local __i = 0

local function make_counter()
  local n = 0
  return {
    inc = function() n = n + 1 end,
    get = function() return n end,
  }
end
local c = make_counter()
c.inc(); c.inc(); c.inc()
do local __t = tostring(c.get()); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
