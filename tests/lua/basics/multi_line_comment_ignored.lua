-- vybe-test: lua/basics/multi_line_comment_ignored
-- origin: languages/lua/tests/lua/test_basics.rs

local __w1 = "comment_ignored"
local __i = 0

--[[
this is a multi-line comment
]]
do local __t = tostring("comment_ignored"); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
