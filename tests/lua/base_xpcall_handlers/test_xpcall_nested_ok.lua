-- vybe-test: lua/base_xpcall_handlers/test_xpcall_nested_ok
-- origin: languages/lua/tests/lua/test_base_xpcall_handlers.rs

local __w1 = "true"
local __i = 0

local function good() return 10 end
local function handler(err) return 0 end
local function outer() return xpcall(good, handler) end
local a, b = pcall(outer)
do local __t = tostring(a == true and b == true and b == true); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
