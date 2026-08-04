-- vybe-test: lua/debug_library_basics/getlocal_val
-- origin: languages/lua/tests/lua/test_debug_library_basics.rs

local __w1 = "x=42"
local __i = 0

local function f()
  local x = 42
  local name, val = debug.getlocal(1, 1)
  do local __t = tostring(name .. "=" .. val); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end
end
f()

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
