-- vybe-test: lua/string_pack_unpack_ext/test_unpack_out_of_bounds
-- origin: languages/lua/tests/lua/test_string_pack_unpack_ext.rs

local __w1 = "false"
local __i = 0

local ok = pcall(function() string.unpack('i4', 'ab') end); do local __t = tostring(tostring(ok)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
