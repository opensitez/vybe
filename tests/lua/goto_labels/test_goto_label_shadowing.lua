-- vybe-test: lua/goto_labels/test_goto_label_shadowing
-- origin: languages/lua/tests/lua/test_goto_labels.rs

local __w1 = "2"
local __i = 0

local a=1; ::lbl::; do ::lbl::; a=2; goto end_lbl; ::end_lbl:: end; do local __t = tostring(a); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
