-- vybe-test: lua/modules_require/test_require_caching
-- origin: languages/lua/tests/lua/test_modules_require.rs

local __w1 = "1"
local __i = 0

local c=0; package.searchers[#package.searchers+1] = function(name) if name=='testmod' then return function() c=c+1; return c end end end; require('testmod'); require('testmod'); do local __t = tostring(c); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
