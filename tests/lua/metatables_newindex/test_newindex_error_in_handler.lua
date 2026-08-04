-- vybe-test: lua/metatables_newindex/test_newindex_error_in_handler
-- origin: languages/lua/tests/lua/test_metatables_newindex.rs

local __w1 = "false"
local __i = 0

local t=setmetatable({}, {__newindex=function() error('boom') end}); local ok, err = pcall(function() t.a=1 end); do local __t = tostring(ok); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
