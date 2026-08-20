-- App 36: Grid Pathfinding (A*)

local W, H = 8, 6
local blocked = { ["3,2"] = true, ["3,3"] = true, ["3,4"] = true, ["5,4"] = true }
local start, goal = {1,1}, {8,6}

local function key(x,y) return x .. "," .. y end
local function h(x,y) return math.abs(goal[1]-x) + math.abs(goal[2]-y) end

local open, g, f, came = {}, {}, {}, {}
open[key(start[1], start[2])] = true
g[key(start[1], start[2])] = 0
f[key(start[1], start[2])] = h(start[1], start[2])

local function best_open()
  local best, bestv
  for k, _ in pairs(open) do
    if not bestv or f[k] < bestv then best, bestv = k, f[k] end
  end
  return best
end

local function parse(k)
  local a,b = k:match("^(%d+),(%d+)$")
  return tonumber(a), tonumber(b)
end

while true do
  local current = best_open()
  if not current then break end
  local cx, cy = parse(current)
  if cx == goal[1] and cy == goal[2] then break end
  open[current] = nil

  local dirs = {{1,0},{-1,0},{0,1},{0,-1}}
  for _, d in ipairs(dirs) do
    local nx, ny = cx + d[1], cy + d[2]
    if nx >= 1 and nx <= W and ny >= 1 and ny <= H and not blocked[key(nx,ny)] then
      local nk = key(nx, ny)
      local tentative = g[current] + 1
      if g[nk] == nil or tentative < g[nk] then
        came[nk] = current
        g[nk] = tentative
        f[nk] = tentative + h(nx, ny)
        open[nk] = true
      end
    end
  end
end

local path, node = {}, key(goal[1], goal[2])
while node do
  path[#path + 1] = node
  node = came[node]
end
for i = #path, 1, -1 do print("step", path[i]) end
