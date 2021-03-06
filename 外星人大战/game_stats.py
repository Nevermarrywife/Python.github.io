class GameStats:
    """跟踪游戏的统计信息"""
    def __init__(self,ai_game):
        """初始化统计信息"""
        self.settings = ai_game.settings
        self.reset_stats()
        #游戏启动时处于活动状态
        self.game_active = False
        #最高分不重置
        self.high_score = 0

    def reset_stats(self):
        """初始化在游戏运行期间可能变化的统计信息"""
        self.score = 0
        self.levle = 1
        self.ships_left = self.settings.ship_limit

