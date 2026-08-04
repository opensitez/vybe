-- vybe-test: lua/error_handling_xpcall_nested/xpcall_handler_fails_nested
-- origin: languages/lua/tests/lua/test_error_handling_xpcall_nested.rs

local __w1 = "true false"
local __i = 0

local function bad_handler(e) error("handler_broke", 0) end
local function outer_handler(e) return "outer:"..e end
local ok, val = xpcall(function()
  return xpcall(function() error("original", 0) end, bad_handler)
end, outer_handler)
do local __t = tostring(ok) .. "\t" .. tostring(val); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
