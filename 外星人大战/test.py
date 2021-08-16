import sys
import pygame

class Test:
    def __init__(self):
        pygame.init()
        self.screen = pygame.display.set_mode((600,400))
        self.screen_rect = self.screen.get_rect()
        pygame.display.set_caption("AI")
        self.bg_color = (230,230,230)
    def run(self):
        while True:
            print(1)
            self.screen.fill(self.bg_color)
            pygame.display.flip()
            #print(dir(self.screen_rect))
if __name__ == '__main__':
        ai = Test()
        ai.run()

