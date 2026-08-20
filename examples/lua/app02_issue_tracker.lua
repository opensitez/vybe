-- Full App 02: Issue Tracker with Labels, Assignees, and SLA report

local issues = {}

local function create_issue(title, severity, assignee, labels)
  local id = #issues + 1
  issues[id] = {
    id = id,
    title = title,
    severity = severity,
    assignee = assignee,
    labels = labels,
    status = "open",
    age_days = 0,
  }
end

local function tick_day()
  for _, i in ipairs(issues) do
    if i.status ~= "closed" then
      i.age_days = i.age_days + 1
    end
  end
end

local function close_issue(id)
  if issues[id] then issues[id].status = "closed" end
end

create_issue("API rate limiter fails under burst", "high", "nora", {"backend", "api"})
create_issue("Docs typo in onboarding", "low", "milo", {"docs"})
create_issue("Payment webhook retries too aggressive", "critical", "nora", {"payments", "sre"})

for _ = 1, 4 do tick_day() end
close_issue(2)
for _ = 1, 3 do tick_day() end

local sla = {low = 14, high = 5, critical = 2}
local by_assignee = {}
for _, i in ipairs(issues) do
  by_assignee[i.assignee] = by_assignee[i.assignee] or {open = 0, breached = 0}
  if i.status == "open" then
    by_assignee[i.assignee].open = by_assignee[i.assignee].open + 1
    if i.age_days > sla[i.severity] then
      by_assignee[i.assignee].breached = by_assignee[i.assignee].breached + 1
    end
  end
  print(string.format("#%d [%s] %s | %s | age=%dd", i.id, i.severity, i.status, i.title, i.age_days))
end

print("assignee summary")
for a, s in pairs(by_assignee) do
  print(a, "open=" .. s.open, "breached=" .. s.breached)
end
