-- vybe-test: lua/globals/global_read_from_deeply_nested_function
-- origin: languages/lua/tests/lua/test_globals.rs

local __w1 = "found"
local __i = 0

deeply_nested = 'found'
local function a()
  local function b()
    local function c() return deeply_nested end
    return c()
  end
  return b()
end
do local __t = tostring(a()); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
