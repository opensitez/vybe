-- vybe-test: lua/xpcall_handler/xpcall_table_error
-- origin: languages/lua/tests/lua/test_xpcall_handler.rs

local __w1 = "false\t7"
local __i = 0

local function handler(e) return e.code end
local ok, v = xpcall(function() error({code=7}) end, handler)
do local __t = tostring(ok) .. "\t" .. tostring(v); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
