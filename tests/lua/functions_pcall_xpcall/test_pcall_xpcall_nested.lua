-- vybe-test: lua/functions_pcall_xpcall/test_pcall_xpcall_nested
-- origin: languages/lua/tests/lua/test_functions_pcall_xpcall.rs

local __w1 = "true false"
local __i = 0

local ok, res = pcall(function() return xpcall(function() error('boom') end, function(e) return 'handled '..e end) end); do local __t = tostring(tostring(ok)..' '..tostring(res)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
