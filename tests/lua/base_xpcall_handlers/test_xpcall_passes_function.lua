-- vybe-test: lua/base_xpcall_handlers/test_xpcall_passes_function
-- origin: languages/lua/tests/lua/test_base_xpcall_handlers.rs

local __w1 = "true"
local __i = 0

local function ok() return 1 end
local function handler(err) return "handled" end
local ok, v = xpcall(ok, handler)
do local __t = tostring(ok and v == 1); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
