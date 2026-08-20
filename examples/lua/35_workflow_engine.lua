-- App 35: Workflow Engine (state + guards + actions)

local flow = {
  submitted = {
    approve = {to = "approved", guard = function(ctx) return ctx.amount < 5000 end},
    reject = {to = "rejected"},
  },
  approved = {
    pay = {to = "paid", guard = function(ctx) return ctx.account_ok end},
  },
}

local function transition(state, event, ctx)
  local edge = flow[state] and flow[state][event]
  if not edge then return state, "invalid" end
  if edge.guard and not edge.guard(ctx) then return state, "guard_failed" end
  return edge.to, "ok"
end

local invoices = {
  {id = 10, amount = 2400, account_ok = true},
  {id = 11, amount = 6400, account_ok = true},
}

for _, inv in ipairs(invoices) do
  local s = "submitted"
  s, msg = transition(s, "approve", inv)
  if msg == "ok" then s, msg = transition(s, "pay", inv) end
  print(string.format("invoice=%d final=%s (%s)", inv.id, s, msg))
end
