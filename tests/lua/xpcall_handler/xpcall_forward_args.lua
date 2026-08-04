-- vybe-test: lua/xpcall_handler/xpcall_forward_args
-- origin: languages/lua/tests/lua/test_xpcall_handler.rs

local __w1 = "true\t42"
local __i = 0

local ok, v = xpcall(function(a, b) return a + b end, function(e) return e end, 10, 32)
do local __t = tostring(ok) .. "\t" .. tostring(v); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
