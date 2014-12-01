
-- for CCLuaEngine
function __G__TRACKBACK__(errorMessage)
    CCLuaLog("----------------------------------------")
    CCLuaLog("LUA ERROR: "..tostring(errorMessage).."\n")
    CCLuaLog(debug.traceback("", 2))
    CCLuaLog("----------------------------------------")
end

CCLuaLoadChunksFromZip("res/zhidianjiangshan.zip")

xpcall(function()    
    require("game")
    game.startup()
end, __G__TRACKBACK__)
