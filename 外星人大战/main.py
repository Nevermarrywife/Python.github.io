import pygame #引入pygame库
import sys
from time import sleep
from game_stats import GameStats
from settings import Settings
from ship import Ship   #从ship实例引入Ship类
from bullet import Bullet
from alien import Alien
from button import Button
from  scoreboard import Scoreboard

class AlienInvasion:
    "管理游戏资源和行为的类"

    "游戏的基本属性"
    def __init__(self):
        # 调用pygame.init()来初始化背景设置，让pygame能正常工作
        pygame.init()
        self.settings = Settings()  #调用Settings()类

        # pygame.display.set_mode()创建显示窗口，宽1200像素ss、高800像素
        self.screen = pygame.display.set_mode((self.settings.screen_width,self.settings.screen_height))

        #创建存储游戏统计信息的实例
        #创建一个用于存储游戏统计信息的实例
        self.stats = GameStats(self)
        self.sb = Scoreboard(self)

        #创建play按钮
        self.play_button = Button(self,"Play")

        # pygame.display.set_caption()设置窗口上的显示内容
        pygame.display.set_caption("Alien Invasion")
        self.ship = Ship(self) #调用Ship实例，self指向的是AlienInvasion实例，可以让Ship访问AlienInvasion的所有对象，如screen等
        self.bullets = pygame.sprite.Group() #存储子弹的编组
        self.aliens = pygame.sprite.Group()#存储飞船的编辑

        self._create_fleet() #运行AlienInvasion时直接调用_create_fleet函数

        #设置背景色
        self.bg_color=(self.settings.bg_color) #给self.bg_color赋上颜色属性

        #self.num = 12 #测试
        #self.tests = Test(self).Print() #测试，Tset(self指向AlienInvasion自己，然后传递到了Test()方法中的xx形参)


    def _create_fleet(self):
        "创建外星人群"

        #创建一个外星人并计算一行可容纳多少个外星人
        #外星人的间距为外星人宽度
        alien = Alien(self)
        alien_width,alien_height = alien.rect.size
        self.aliens.add(alien)
        alien_width = alien.rect.width
        available_space_x = self.settings.screen_width - (2 * alien_width)
        number_aliens_x = available_space_x // (2 * alien_width)

        # 计算屏幕可容纳多少行外星人
        ship_height = self.ship.rect.height
        available_space_y = (self.settings.screen_height -
                             (3 * alien_height) - ship_height)
        number_rows = available_space_y // (2 * alien_height)

        #创建外星人群
        for row_number in range(number_rows):
            for alien_number in range(number_aliens_x):
                self._creat_alien(alien_number,row_number)

    def _creat_alien(self,alien_number,row_number):
        """创建一个外星人并将其加入当前行"""
        alien = Alien(self)
        alien_width,alien_height = alien.rect.size
        alien.x = alien_width + 2 * alien_width * alien_number
        alien.rect.x = alien.x
        alien.rect.y = alien.rect.height + 2 * alien.rect.height * row_number
        self.aliens.add(alien)

    def run_game(self):
        "开始游戏的主循环"

        #self.ship.test() 测试
        #print(dir(self))
        self.test()
        while True:
            self._check_events() #调用事件管理方法

            if self.stats.game_active:
                self.ship.update() #移动飞船
                self._update_bullets()
                self._update_aliens()

            self._update_screen()  # 更新屏幕
            #print(len(self.bullets))    #核实还有多少颗子弹
            pygame.display.flip() #让最近绘制的屏幕可见。

    def _update_aliens(self):
        #更新所有外星人群的位置
        self.aliens.update()
        #检查是否外星人触碰屏幕边缘并且改变移动位置
        self._check_fleet_edges()

        #检测外星人和飞船之间的碰撞
        if pygame.sprite.spritecollideany(self.ship,self.aliens):
            self._ship_hit()

        #检查是否有外星人到达屏幕底端
        self._check_aliens_bottom()

    def _ship_hit(self):
        """响应飞船被外星人撞到"""

        if self.stats.ships_left > 1:
            #将ships_left减1
            self.stats.ships_left -= 1
            self.sb.prep_ships()


            #清空余下的外星人和子弹
            self.aliens.empty()
            self.bullets.empty()

            #创建一群新的外星人，并将飞船放在屏幕底端的中央
            self._create_fleet()
            self.ship.center_ship()

            #暂停
            sleep(0.5)


        else:
            self.stats.game_active = False
            pygame.mouse.set_visible(True)

    def _check_fleet_edges(self):
        """有外星人到达边缘时采取相应的措施"""

        for alien in self.aliens.sprites():
            if alien.check_edges():
                self._change_fleet_direction()
                break

    def _change_fleet_direction(self):
        """整群外星人下移，并改变方向"""

        for alien in self.aliens.sprites():
            alien.rect.y += self.settings.fleet_drop_speed
        self.settings.fleet_direction *= -1

    def _check_aliens_bottom(self):
        """检查是否有外星人到达了屏幕底端"""

        screen_rect = self.screen.get_rect()
        for alien in self.aliens.sprites():
            if alien.rect.bottom >= screen_rect.bottom:
                #像飞船撞到一样处理
                self._ship_hit()
                break

    def _check_events(self):
        """监视键盘和鼠标事件"""

        for event in pygame.event.get():
            if event.type == pygame.QUIT:
                sys.exit()

            elif event.type == pygame.KEYDOWN: #案件按下,调用_check_keydown_events函数
                self._check_keydown_events(event)

            elif event.type == pygame.KEYUP: #按键松起
                self._check_keyup_events(event)

            elif event.type == pygame.MOUSEBUTTONDOWN:
                mouse_pos = pygame.mouse.get_pos() #鼠标点击时的坐标
                self._check_play_button(mouse_pos)

    def _check_play_button(self,mouse_pos):
        """玩家点击play时开始游戏"""
        button_clicked = self.play_button.rect.collidepoint(mouse_pos)
        if button_clicked and not self.stats.game_active:
            #重置游戏设置
            self.settings.initialize_dynamic_settings()
            #重置游戏统计信息
            self.stats.reset_stats()
            self.stats.game_active = True
            self.sb.prep_score()
            self.sb.prep_level()
            self.sb.prep_ships()

            #清空余下的外星人和子弹
            self.aliens.empty()
            self.bullets.empty()

            #创建一群新的外星人并让飞船居中
            self._create_fleet()
            self.ship.center_ship()

            #隐藏鼠标光标
            pygame.mouse.set_visible(False)

    def _check_keydown_events(self,event):
        "相应按键"

        if event.key == pygame.K_RIGHT:
            self.ship.moving_right = True
        elif event.key == pygame.K_LEFT:
            self.ship.moving_lift = True
        elif event.key == pygame.K_q:
            sys.exit()  #退出游戏
        elif event.key == pygame.K_SPACE:  # 发射子弹
            self._fire_bullet()

    def _fire_bullet(self):
        "创建一颗子弹，并将其加入编组bullets中"

        if len(self.bullets) < self.settings.bullets_allowed:
            new_bullet = Bullet(self)
            self.bullets.add(new_bullet)

    def _check_keyup_events(self,event):
        "响应松起"

        if event.key == pygame.K_RIGHT:
            self.ship.moving_right = False
        elif event.key == pygame.K_LEFT:
            self.ship.moving_lift = False

    def _update_bullets(self):
        "更新子弹的位置并删除消失的子弹"

        # 更新子弹的位置
        self.bullets.update()
        # 删除消失的子弹
        for bullet in self.bullets.copy():
            if bullet.rect.bottom <= 0:
                self.bullets.remove(bullet)
        self._check_bullet_alien_collisions()

    def _check_bullet_alien_collisions(self):
        """检查是否有子弹击中了外星人"""

        # 如果是，删除相应子弹和外星人
        collisions = pygame.sprite.groupcollide(
            self.bullets,self.aliens,False,True
        )
        if collisions:
            for aliens in collisions.values():
                self.stats.score += self.settings.alien_points
                self.sb.prep_score()
                self.sb.check_high_score()
        if not self.aliens:
            #删除现有的子弹并新建一群外星人
            self.bullets.empty()
            self._create_fleet()
            self.settings.increase_speed()

            #提高等级
            self.stats.levle +=1
            self.sb.prep_level()

    def _update_screen(self):
        """run_game每次循环都调用该函数，进行重新绘制屏幕"""

        self.screen.fill(self.bg_color)
        self.ship.blitme() #每次循环绘制飞船
        for bullet in self.bullets.sprites():
            bullet.draw_bullet()
        self.aliens.draw(self.screen)
        self.sb.show_score()

        #如果游戏处于非活动状态，就绘制Playanniu
        if not self.stats.game_active:
            self.play_button.draw_button()

    def test(self):
        """测试用"""
        #print(self.ship.rect.x)
        #print(self.aliens)
        True

if __name__ == '__main__':
    #创建游戏实例并运行游戏
    ai = AlienInvasion()
    ai.run_game()

