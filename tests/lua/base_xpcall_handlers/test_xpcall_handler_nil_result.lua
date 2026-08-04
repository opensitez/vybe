-- vybe-test: lua/base_xpcall_handlers/test_xpcall_handler_nil_result
-- origin: languages/lua/tests/lua/test_base_xpcall_handlers.rs

local __w1 = "true"
local __i = 0

local function bad() error("x") end
local function handler(err) return nil end
local ok, v = xpcall(bad, handler)
do local __t = tostring(ok == false and v == nil); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
