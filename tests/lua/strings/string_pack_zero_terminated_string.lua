-- vybe-test: lua/strings/string_pack_zero_terminated_string
-- origin: languages/lua/tests/lua/test_strings.rs

local __w1 = "4 abc"
local __i = 0

local s = string.pack("z", "abc")
do local __t = tostring(#s .. " " .. string.unpack("z", s)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
