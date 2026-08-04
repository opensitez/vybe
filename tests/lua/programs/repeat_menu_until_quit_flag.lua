-- vybe-test: lua/programs/repeat_menu_until_quit_flag
-- origin: languages/lua/tests/lua/test_programs.rs

local __w1 = "2"
local __i = 0

local choice = 0
local guard = 0
repeat
  guard = guard + 1
  choice = guard
until choice >= 2
do local __t = tostring(choice); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
