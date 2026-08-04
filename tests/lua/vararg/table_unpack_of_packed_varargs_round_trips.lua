-- vybe-test: lua/vararg/table_unpack_of_packed_varargs_round_trips
-- origin: languages/lua/tests/lua/test_vararg.rs

local __w1 = "7,8,9"
local __i = 0

local function pack_and_unpack(...)
  local t = table.pack(...)
  return table.unpack(t, 1, t.n)
end
local a, b, c = pack_and_unpack(7, 8, 9)
do local __t = tostring(a .. ',' .. b .. ',' .. c); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
