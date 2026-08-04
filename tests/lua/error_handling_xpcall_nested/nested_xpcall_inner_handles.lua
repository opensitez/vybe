-- vybe-test: lua/error_handling_xpcall_nested/nested_xpcall_inner_handles
-- origin: languages/lua/tests/lua/test_error_handling_xpcall_nested.rs

local __w1 = "true false true false"
local __i = 0

local inner_run = false
local outer_run = false
local function inner_handler(e) inner_run = true; return "inner:"..e end
local function outer_handler(e) outer_run = true; return "outer:"..e end
local ok, val = xpcall(function()
  return xpcall(function() error("fail", 0) end, inner_handler)
end, outer_handler)
do local __t = tostring(ok) .. "\t" .. tostring(val) .. "\t" .. tostring(inner_run) .. "\t" .. tostring(outer_run); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
