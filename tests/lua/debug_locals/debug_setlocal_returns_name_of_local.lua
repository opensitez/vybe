-- vybe-test: lua/debug_locals/debug_setlocal_returns_name_of_local
-- origin: languages/lua/tests/lua/test_debug_locals.rs

local __w1 = "x 20"
local __i = 0

local function f()
  local x = 10
  local name = debug.setlocal(1, 1, 20)
  return name .. " " .. x
end
do local __t = tostring(f()); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
