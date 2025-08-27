import sys
import os
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QHBoxLayout, QPushButton, QTextEdit, QFileDialog, 
                             QLabel, QLineEdit, QListWidget, QGroupBox, 
                             QTabWidget, QCheckBox, QSpinBox, QComboBox,
                             QMessageBox, QProgressBar, QFormLayout, QSizePolicy)
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QTimer, QRectF, QPointF
from PyQt5.QtGui import QFont, QPainter, QColor, QPen
import pandas as pd
import math

class ParticleAnimation(QWidget):
    """粒子动画组件"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.particles = []
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.update_animation)
        self.timer.start(30)  # 30ms更新一次
        
        # 初始化粒子
        for i in range(500):
            self.particles.append({
                'x': float(i * 40 % self.width()),
                'y': float(self.height() - (i * 30) % self.height()),
                'size': 2 + (i % 5),
                'speed': 0.5 + (i % 3) * 0.5,
                'color': QColor(50 + (i * 5) % 200, 100 + (i * 3) % 150, 150 + (i * 7) % 100),
                'angle': float(i % 360)
            })
    
    def update_animation(self):
        for particle in self.particles:
            # 更新粒子位置
            particle['x'] += math.cos(math.radians(particle['angle'])) * particle['speed']
            particle['y'] += math.sin(math.radians(particle['angle'])) * particle['speed']
            
            # 更新角度
            particle['angle'] += 0.5
            
            # 边界检查
            if particle['x'] > self.width():
                particle['x'] = 0
            elif particle['x'] < 0:
                particle['x'] = self.width()
                
            if particle['y'] > self.height():
                particle['y'] = 0
            elif particle['y'] < 0:
                particle['y'] = self.height()
        
        self.update()
    
    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)
        
        # 绘制粒子
        for i, particle in enumerate(self.particles):
            painter.setPen(QPen(particle['color'], particle['size']))
            painter.drawPoint(int(particle['x']), int(particle['y']))
            
            # 绘制连线（与附近的粒子）
            if i < len(self.particles) - 1:
                next_particle = self.particles[i + 1]
                distance = math.sqrt((particle['x'] - next_particle['x']) ** 2 + 
                                   (particle['y'] - next_particle['y']) ** 2)
                if distance < 100:  # 只连接距离较近的粒子
                    painter.setPen(QPen(particle['color'], 1))
                    painter.drawLine(int(particle['x']), int(particle['y']), 
                                   int(next_particle['x']), int(next_particle['y']))
    
    def resizeEvent(self, event):
        # 确保粒子在窗口大小改变时仍然在可视范围内
        for particle in self.particles:
            if particle['x'] > self.width():
                particle['x'] = self.width()
            if particle['y'] > self.height():
                particle['y'] = self.height()