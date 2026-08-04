-- vybe-test: lua/metatables_newindex/test_newindex_function_modifies_other_table
-- origin: languages/lua/tests/lua/test_metatables_newindex.rs

local __w1 = "nil 5"
local __i = 0

local other={}; local t={}; setmetatable(t, {__newindex=function(tbl, k, v) other[k]=v end}); t.a=5; do local __t = tostring((t.a or 'nil')..' '..other.a); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
