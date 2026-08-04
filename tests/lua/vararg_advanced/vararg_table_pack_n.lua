-- vybe-test: lua/vararg_advanced/vararg_table_pack_n
-- origin: languages/lua/tests/lua/test_vararg_advanced.rs

local __w1 = "3"
local __i = 0

local function f(...)
  local t = table.pack(...)
  return t.n
end
do local __t = tostring(f(5, nil, 7)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
