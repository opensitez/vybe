-- Vybe test harness — Lua.
--
-- Real Lua: this runs under `lua` unchanged. Like the C, COBOL and Fortran
-- harnesses it documents the shape the emitter produces rather than providing
-- a function to splice, because the check has to be at the CALL SITE.
--
-- WHY NOT COLLECT THE OUTPUT.
-- Every other collecting harness (ruby, java, dart, pascal, php) intercepts
-- the language's print. Lua cannot, under Vybe — measured, all three forms
-- fall through to the real stdout:
--
--   function print(...) ... end     -- builtin still used
--   print = function(...) ... end   -- builtin still used
--   _G.print = function(...) ... end -- builtin still used
--
-- (`io.write` is also undefined.) So each `print` is compared where it stands.
--
-- WHY ONLY THE FIRST PRINT.
-- The corpus asserts through `run_lua_one`, which takes `run_lua(src)` and
-- returns only its FIRST line. A program that prints more is not asserting
-- those lines, so checking them would invent expectations the test never had.
-- The counter still advances on every print, which is what lets a print inside
-- a LOOP be handled: the check fires on print #1 wherever it occurs.
--
-- The emitted shape, for `print(a .. "," .. b)` expecting "2,1":
--
--   local __w1 = "2,1"
--   local __i = 0
--   ...
--   do local __t = tostring(a .. "," .. b); __i = __i + 1
--      if __i == 1 and __t ~= __w1 then
--        error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end
--   ...
--   if __i == 0 then error("FAIL: no output") end
--
-- Multiple arguments join with a TAB, which is what Lua's `print` does —
-- verified identical in both runtimes.
--
-- Failure is `error(...)`, which exits non-zero in both (lua 1, Vybe 1) and
-- carries its message, so the expected and actual values survive.

local __w1 = "2,1"
local __i = 0

local a, b = 1, 2
a, b = b, a

do
  local __t = tostring(a .. "," .. b)
  __i = __i + 1
  if __i == 1 and __t ~= __w1 then
    error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]")
  end
end

if __i == 0 then error("FAIL: no output") end
