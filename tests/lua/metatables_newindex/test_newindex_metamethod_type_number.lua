-- vybe-test: lua/metatables_newindex/test_newindex_metamethod_type_number
-- origin: languages/lua/tests/lua/test_metatables_newindex.rs

local __w1 = "42"
local __i = 0

local out; debug.setmetatable(0, {__newindex=function(n, k, v) out=v end}); (10).foo=42; do local __t = tostring(out); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
