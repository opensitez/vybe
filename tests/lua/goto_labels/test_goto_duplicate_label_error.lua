-- vybe-test: lua/goto_labels/test_goto_duplicate_label_error
-- origin: languages/lua/tests/lua/test_goto_labels.rs

local __w1 = "false"
local __i = 0

local ok = pcall(function() load('::lbl::; ::lbl::') end); do local __t = tostring(tostring(ok)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
