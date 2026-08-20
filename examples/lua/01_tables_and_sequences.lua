-- App 01: Task Board CLI (projects, priorities, tags, archive)

local board = {
  projects = {
    alpha = {
      {id = 1, title = "Design schema", priority = 3, tags = {"db", "design"}, done = false},
      {id = 2, title = "Build import tool", priority = 2, tags = {"etl"}, done = true},
    },
    beta = {
      {id = 3, title = "Write docs", priority = 1, tags = {"docs"}, done = false},
    }
  },
  archive = {}
}

local function has_tag(item, wanted)
  for _, t in ipairs(item.tags) do
    if t == wanted then return true end
  end
  return false
end

local function list_open(project, tag)
  local items = board.projects[project] or {}
  table.sort(items, function(a, b) return a.priority > b.priority end)
  print("Open tasks for project:", project)
  for _, item in ipairs(items) do
    if not item.done and (not tag or has_tag(item, tag)) then
      print(string.format("[#%d] p%d %s (%s)", item.id, item.priority, item.title, table.concat(item.tags, ",")))
    end
  end
end

local function complete_task(project, id)
  for _, item in ipairs(board.projects[project] or {}) do
    if item.id == id and not item.done then
      item.done = true
      board.archive[#board.archive + 1] = {project = project, id = item.id, title = item.title}
      return true
    end
  end
  return false
end

list_open("alpha")
print("mark done:", complete_task("alpha", 1))
list_open("alpha")
print("archive entries:", #board.archive)
