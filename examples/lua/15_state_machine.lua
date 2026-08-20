-- 15_state_machine.lua
-- Demonstrates table-driven finite state machine style.

local transitions = {
  idle = {
    coin = "ready"
  },
  ready = {
    select_item = "dispense",
    refund = "idle"
  },
  dispense = {
    complete = "idle"
  }
}

local state = "idle"
local events = {"coin", "select_item", "complete", "coin", "refund"}

local function step(current, event)
  local next_state = transitions[current] and transitions[current][event]
  if not next_state then
    return current, "ignored"
  end
  return next_state, "ok"
end

for i = 1, #events do
  local event = events[i]
  local prev = state
  state, status = step(state, event)
  print(string.format("%d) %s --(%s/%s)--> %s", i, prev, event, status, state))
end
