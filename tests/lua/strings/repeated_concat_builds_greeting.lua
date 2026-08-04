-- vybe-test: lua/strings/repeated_concat_builds_greeting
-- origin: languages/lua/tests/lua/test_strings.rs

local __w1 = "hello lua!"
local __i = 0

local name = "lua"
do local __t = tostring("hello " .. name .. "!"); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
