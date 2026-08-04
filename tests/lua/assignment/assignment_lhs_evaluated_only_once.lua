-- vybe-test: lua/assignment/assignment_lhs_evaluated_only_once
-- origin: languages/lua/tests/lua/test_assignment.rs

local __w1 = "TK=42"
local __i = 0

local order = ""
local t = {}
local function get_t() order = order .. "T"; return t end
local function get_k() order = order .. "K"; return "key" end
get_t()[get_k()] = 42
do local __t = tostring(order .. "=" .. t.key); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
