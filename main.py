import sys
import json
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
                             QLabel, QLineEdit, QPushButton, QTableWidget, QTableWidgetItem,
                             QDateEdit, QComboBox, QMessageBox, QTabWidget, QHeaderView,
                             QFileDialog, QInputDialog, QFrame, QDialog, QGridLayout,
                             QProgressBar)
from PyQt5.QtCore import QDate, Qt, QThread, pyqtSignal, pyqtSlot
import requests
import pandas as pd
from io import StringIO
from datetime import datetime, timedelta


class BatchAddDatesDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("批量添加交易日")
        self.setup_ui()

    def setup_ui(self):
        layout = QGridLayout()

        layout.addWidget(QLabel("起始日期:"), 0, 0)
        self.start_date_edit = QDateEdit()
        self.start_date_edit.setCalendarPopup(True)
        self.start_date_edit.setDate(QDate.currentDate())
        layout.addWidget(self.start_date_edit, 0, 1)

        layout.addWidget(QLabel("交易日数量:"), 1, 0)
        self.days_count_edit = QLineEdit()
        self.days_count_edit.setText("20")
        layout.addWidget(self.days_count_edit, 1, 1)

        layout.addWidget(QLabel("跳过周末:"), 2, 0)
        self.skip_weekend_check = QComboBox()
        self.skip_weekend_check.addItems(["是", "否"])
        layout.addWidget(self.skip_weekend_check, 2, 1)

        self.ok_btn = QPushButton("确定")
        self.ok_btn.clicked.connect(self.accept)
        layout.addWidget(self.ok_btn, 3, 0, 1, 2)

        self.setLayout(layout)

    def get_dates(self):
        start_date = self.start_date_edit.date().toPyDate()
        days_count = int(self.days_count_edit.text())
        skip_weekend = self.skip_weekend_check.currentText() == "是"

        dates = []
        current_date = start_date
        while len(dates) < days_count:
            if skip_weekend and current_date.weekday() >= 5:  # 周六或周日
                current_date += timedelta(days=1)
                continue
            dates.append(current_date.strftime("%Y-%m-%d"))
            current_date += timedelta(days=1)
        return dates


class DataRefreshThread(QThread):
    """用于刷新市场数据的线程，避免UI卡顿"""
    progress_updated = pyqtSignal(int, str)
    finished = pyqtSignal(bool, str)

    def __init__(self, parent, option_name, query_date, keyword=None, na_dates=None):
        super().__init__(parent)
        self.parent = parent
        self.option_name = option_name
        self.query_date = query_date
        self.keyword = keyword  # 用于筛选需要刷新的数据
        self.na_dates = na_dates if na_dates is not None else {}  # 格式: {期权名称: [日期列表]}
        self.total_tasks = 0
        self.completed_tasks = 0
        self.is_canceled = False

    def run(self):
        try:
            # 筛选需要处理的期权
            target_options = self.parent.options
            if self.keyword:
                target_options = {
                    name: option for name, option in self.parent.options.items()
                    if self.keyword.lower() in option["name"].lower()
                }
                if not target_options:
                    self.finished.emit(False, f"没有找到包含关键词 '{self.keyword}' 的期权数据!")
                    return

            # 计算总任务量
            if self.na_dates:  # 只刷新N/A数据
                for name, dates in self.na_dates.items():
                    if name in target_options:  # 只处理筛选后的期权
                        self.total_tasks += len(dates)

                # 执行刷新
                for name, dates in self.na_dates.items():
                    if self.is_canceled or name not in target_options:
                        break

                    option = self.parent.options[name]

                    for date in dates:
                        if self.is_canceled:
                            break

                        self.progress_updated.emit(
                            int(self.completed_tasks / self.total_tasks * 100) if self.total_tasks > 0 else 0,
                            f"正在重新获取 {option['name']} 在 {date} 的数据..."
                        )
                        self.parent.refresh_option_data(option, date)
                        self.completed_tasks += 1
                        self.progress_updated.emit(
                            int(self.completed_tasks / self.total_tasks * 100) if self.total_tasks > 0 else 100,
                            f"已处理 {option['name']} 在 {date} 的数据"
                        )
            elif self.option_name is None:  # 所有筛选后的期权
                # 计算总任务量
                for name, option in target_options.items():
                    for date in option["trade_dates"]:
                        if date <= self.query_date:
                            self.total_tasks += 1

                # 执行刷新
                for name, option in target_options.items():
                    if self.is_canceled:
                        break

                    dates_to_update = []
                    for date in option["trade_dates"]:
                        if date <= self.query_date:
                            dates_to_update.append(date)

                    for date in dates_to_update:
                        if self.is_canceled:
                            break

                        self.progress_updated.emit(
                            int(self.completed_tasks / self.total_tasks * 100) if self.total_tasks > 0 else 0,
                            f"正在获取 {option['name']} 在 {date} 的数据..."
                        )
                        self.parent.refresh_option_data(option, date)
                        self.completed_tasks += 1
                        self.progress_updated.emit(
                            int(self.completed_tasks / self.total_tasks * 100) if self.total_tasks > 0 else 100,
                            f"已获取 {option['name']} 在 {date} 的数据"
                        )
            else:  # 特定期权
                if self.option_name in target_options:
                    option = self.parent.options[self.option_name]
                    # 计算总任务量
                    for date in option["trade_dates"]:
                        if date <= self.query_date:
                            self.total_tasks += 1

                    # 执行刷新
                    for date in option["trade_dates"]:
                        if self.is_canceled or date > self.query_date:
                            break

                        self.progress_updated.emit(
                            int(self.completed_tasks / self.total_tasks * 100) if self.total_tasks > 0 else 0,
                            f"正在获取 {option['name']} 在 {date} 的数据..."
                        )
                        self.parent.refresh_option_data(option, date)
                        self.completed_tasks += 1
                        self.progress_updated.emit(
                            int(self.completed_tasks / self.total_tasks * 100) if self.total_tasks > 0 else 100,
                            f"已获取 {option['name']} 在 {date} 的数据"
                        )

            if self.is_canceled:
                self.finished.emit(False, "操作已取消")
                return

            # 保存更新后的数据
            self.parent.save_data()
            self.finished.emit(True, "市场数据已重新获取并更新!")
        except Exception as e:
            self.finished.emit(False, f"发生错误: {str(e)}")

    def cancel(self):
        self.is_canceled = True


class QueryThread(QThread):
    """用于执行查询操作的线程，避免UI卡顿"""
    progress_updated = pyqtSignal(int, str)
    finished = pyqtSignal(bool, dict)  # 第二个参数是错误信息字典 {期权名称: [日期列表]}
    result_ready = pyqtSignal(dict)  # 用于传递查询结果

    def __init__(self, parent, query_date, option_name=None, keyword=None, is_keyword_query=False):
        super().__init__(parent)
        self.parent = parent
        self.query_date = query_date
        self.option_name = option_name
        self.keyword = keyword  # 查询关键词
        self.is_keyword_query = is_keyword_query  # 标识是否是关键词查询
        self.is_canceled = False
        self.results = {
            "single_option": [],
            "active_options": [],
            "expired_options": []
        }
        self.error_messages = {}

    def run(self):
        try:
            if not self.parent.options:
                self.finished.emit(False, {"error": "没有可查询的期权数据!"})
                return

            # 保存当前数据
            self.parent.save_data()

            # 处理关键词筛选
            filtered_options = self.parent.options.copy()
            if self.is_keyword_query and self.keyword:
                filtered_options = {
                    name: option for name, option in self.parent.options.items()
                    if self.keyword.lower() in option["name"].lower()
                }

                if not filtered_options:
                    self.finished.emit(False, {"error": f"没有找到包含关键词 '{self.keyword}' 的期权数据!"})
                    return

            # 第一阶段：计算需要查询的数据量并检查N/A
            na_dates = {}  # 存储需要重新获取的日期 {期权名称: [日期列表]}
            total_options = len(filtered_options) if self.option_name is None and not self.is_keyword_query else 1
            if self.is_keyword_query:
                total_options = len(filtered_options)

            processed_options = 0

            self.progress_updated.emit(0, "准备查询数据...")

            # 确定要查询的期权集合
            target_options = filtered_options
            if not self.is_keyword_query and self.option_name and self.option_name in self.parent.options:
                target_options = {self.option_name: self.parent.options[self.option_name]}

            if not target_options:
                self.finished.emit(False, {"error": "未找到目标期权数据!"})
                return

            for name, option in target_options.items():
                if self.is_canceled:
                    break

                processed_options += 1
                progress_percent = int(processed_options / total_options * 30)  # 第一阶段占30%进度
                self.progress_updated.emit(progress_percent, f"检查 {option['name']} 的数据... ({progress_percent}%)")

                na_dates[name] = []
                for date in option["trade_dates"]:
                    # 修改处1：原代码是检查所有查询日期前的N/A数据
                    # 原代码：if date < self.query_date and option["close_prices"].get(date) == "N/A":
                    # 修改为：只检查查询日期当天的N/A数据（如果是查询当天）和查询日期前的N/A数据
                    if (date == self.query_date or date < self.query_date) and option["close_prices"].get(date) == "N/A":
                        na_dates[name].append(date)

                # 如果没有需要重新获取的日期，移除该期权
                if not na_dates[name]:
                    del na_dates[name]

            # 第二阶段：重新获取N/A数据
            na_refresh_success = True
            if na_dates and not self.is_canceled:
                # 计算总任务量
                total_na_tasks = sum(len(dates) for dates in na_dates.values())
                completed_na_tasks = 0

                if total_na_tasks > 0:
                    for name, dates in na_dates.items():
                        if self.is_canceled:
                            break

                        if name not in self.parent.options:
                            continue

                        option = self.parent.options[name]

                        for date in dates:
                            if self.is_canceled:
                                break

                            progress_percent = 30 + int(completed_na_tasks / total_na_tasks * 40)  # 第二阶段占30-70%进度
                            self.progress_updated.emit(progress_percent,
                                                       f"正在获取 {option['name']} 在 {date} 的数据... ({progress_percent}%)")

                            # 修改处2：原代码无条件刷新N/A数据
                            # 修改为：如果是查询日期当天且是N/A，尝试刷新但不报错；查询日期前仍报错
                            if date == self.query_date:
                                # 查询日期当天，尝试刷新但不报错
                                self.parent.refresh_option_data(option, date)
                                if option["close_prices"].get(date) == "N/A":
                                    # 当天数据仍为N/A，不报错
                                    pass
                            else:
                                # 查询日期前，保持原逻辑
                                self.parent.refresh_option_data(option, date)
                                if option["close_prices"].get(date) == "N/A":
                                    if name not in self.error_messages:
                                        self.error_messages[name] = []
                                    self.error_messages[name].append(date)

                            completed_na_tasks += 1
                            self.progress_updated.emit(
                                int(completed_na_tasks / total_na_tasks * 100) if total_na_tasks > 0 else 100,
                                f"已处理 {option['name']} 在 {date} 的数据"
                            )

            # 第三阶段：计算并准备查询结果
            if not self.is_canceled and na_refresh_success:
                current_date = QDate.currentDate().toString("yyyy-MM-dd")
                active_count = 0
                expired_count = 0
                target_count = len(target_options)

                # 根据结果数量决定查询模式
                use_single_mode = target_count == 1

                if not use_single_mode:  # 多期权查询模式
                    total_options = len(target_options)
                    processed_options = 0

                    for name, option in target_options.items():
                        if self.is_canceled:
                            break

                        processed_options += 1
                        progress_percent = 70 + int(processed_options / total_options * 30)  # 第三阶段占70-100%进度
                        self.progress_updated.emit(progress_percent,
                                                   f"处理 {option['name']} 的数据... ({progress_percent}%)")

                        last_trade_date = max(option["trade_dates"]) if option["trade_dates"] else ""

                        # 判断期权是否已到期
                        if last_trade_date and last_trade_date < current_date:  # 已到期期权
                            # 使用最后交易日的数据
                            if last_trade_date in option["trade_dates"]:
                                self.parent.calculate_option_data(option, last_trade_date)
                                self.results["expired_options"].append({
                                    "date": last_trade_date,
                                    "option": option
                                })
                                expired_count += 1
                        elif last_trade_date:  # 未到期期权
                            # 使用查询日期的数据
                            if self.query_date in option["trade_dates"]:
                                self.parent.calculate_option_data(option, self.query_date)
                                self.results["active_options"].append({
                                    "date": self.query_date,
                                    "option": option
                                })
                                active_count += 1

                    self.results["active_count"] = active_count
                    self.results["expired_count"] = expired_count
                else:  # 单个期权查询模式
                    # 获取唯一的期权
                    name, option = next(iter(target_options.items()))
                    self.progress_updated.emit(70, f"处理 {option['name']} 的数据... (70%)")

                    # 显示该期权所有交易日数据（截止查询日期）
                    valid_dates = [date for date in option["trade_dates"] if date <= self.query_date]
                    total_dates = len(valid_dates)

                    for i, date in enumerate(valid_dates):
                        if self.is_canceled:
                            break

                        self.parent.calculate_option_data(option, date)
                        self.results["single_option"].append({
                            "date": date,
                            "option": option
                        })

                        progress_percent = 70 + int((i + 1) / total_dates * 30)  # 第三阶段占70-100%进度
                        self.progress_updated.emit(progress_percent,
                                                   f"处理 {option['name']} 在 {date} 的数据... ({progress_percent}%)")

                # 收集错误信息
                for name, option in target_options.items():
                    for date in option["trade_dates"]:
                        if date < self.query_date and option["close_prices"].get(date) == "N/A":
                            if name not in self.error_messages:
                                self.error_messages[name] = []
                            self.error_messages[name].append(date)

                self.result_ready.emit(self.results)
                self.progress_updated.emit(100, "查询完成")
                self.finished.emit(True, self.error_messages)
            elif self.is_canceled:
                self.finished.emit(False, {"error": "查询已取消"})
            else:
                self.finished.emit(False, {"error": "N/A数据刷新失败"})
        except Exception as e:
            self.finished.emit(False, {"error": f"发生错误: {str(e)}"})

    def cancel(self):
        self.is_canceled = True


class OptionPositionCalculator(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("期权头寸计算及期货数据统计系统")
        self.setGeometry(100, 100, 1200, 800)

        self.options = {}  # 存储所有期权数据
        self.current_option = None
        self.data_file = "options_data.json"  # 默认数据文件名
        self.refresh_thread = None
        self.query_thread = None  # 查询线程
        self.query_in_progress = False
        self.na_refresh_thread = None  # 用于N/A数据刷新的线程

        self.init_ui()
        self.load_data()  # 尝试加载保存的数据

    def init_ui(self):
        # 创建主控件和布局
        main_widget = QWidget()
        main_layout = QVBoxLayout()

        # 创建菜单栏
        self.create_menu_bar()

        # 创建标签页
        self.tab_widget = QTabWidget()

        # 添加期权录入标签页
        input_tab = QWidget()
        self.setup_input_tab(input_tab)
        self.tab_widget.addTab(input_tab, "期权录入/修改")

        # 添加查询标签页
        query_tab = QWidget()
        self.setup_query_tab(query_tab)
        self.tab_widget.addTab(query_tab, "数据查询")

        # 添加平仓操作标签页
        close_tab = QWidget()
        self.setup_close_tab(close_tab)
        self.tab_widget.addTab(close_tab, "平仓操作")

        main_layout.addWidget(self.tab_widget)

        # 添加进度显示区域
        progress_container = QWidget()
        progress_layout = QVBoxLayout(progress_container)
        progress_layout.setContentsMargins(5, 5, 5, 5)

        # 进度信息标签
        self.progress_label = QLabel("准备就绪")
        self.progress_label.setStyleSheet("color: #666; font-family: SimSun; font-size: 16px;")
        progress_layout.addWidget(self.progress_label)

        # 进度条
        self.progress_bar = QProgressBar()
        self.progress_bar.setMinimum(0)
        self.progress_bar.setMaximum(100)
        self.progress_bar.setValue(0)
        self.progress_bar.setTextVisible(True)
        self.progress_bar.setStyleSheet("QProgressBar { font-family: SimSun; font-size: 16px; }")
        self.progress_bar.hide()  # 初始隐藏
        progress_layout.addWidget(self.progress_bar)

        # 取消按钮
        self.cancel_btn = QPushButton("取消")
        self.cancel_btn.setStyleSheet("font-family: SimSun; font-size: 16px;")
        self.cancel_btn.clicked.connect(self.cancel_operation)
        self.cancel_btn.hide()  # 初始隐藏
        progress_layout.addWidget(self.cancel_btn)

        main_layout.addWidget(progress_container)

        main_widget.setLayout(main_layout)
        self.setCentralWidget(main_widget)

    def create_menu_bar(self):
        menubar = self.menuBar()

        # 文件菜单
        file_menu = menubar.addMenu('文件')

        save_action = file_menu.addAction('保存数据')
        save_action.triggered.connect(self.save_data)

        save_as_action = file_menu.addAction('另存为...')
        save_as_action.triggered.connect(self.save_data_as)

        load_action = file_menu.addAction('加载数据...')
        load_action.triggered.connect(self.load_data_from_file)

        file_menu.addSeparator()

        exit_action = file_menu.addAction('退出')
        exit_action.triggered.connect(self.close)

    def setup_input_tab(self, tab):
        layout = QVBoxLayout()

        # 期权选择
        select_group = QWidget()
        select_layout = QHBoxLayout()

        select_layout.addWidget(QLabel("选择期权:"))
        self.option_select_combo = QComboBox()
        self.option_select_combo.currentIndexChanged.connect(self.load_option_for_edit)
        select_layout.addWidget(self.option_select_combo)

        select_layout.addStretch()

        select_group.setLayout(select_layout)
        layout.addWidget(select_group)

        # 期权基本信息输入
        info_group = QWidget()
        info_layout = QHBoxLayout()

        info_layout.addWidget(QLabel("期权名称:"))
        self.option_name_input = QLineEdit()
        info_layout.addWidget(self.option_name_input)

        info_layout.addWidget(QLabel("期货代码:"))
        self.option_code_input = QLineEdit()
        info_layout.addWidget(self.option_code_input)

        info_layout.addWidget(QLabel("执行价格:"))
        self.strike_price_input = QLineEdit()
        info_layout.addWidget(self.strike_price_input)

        info_layout.addWidget(QLabel("初始计提量:"))
        self.initial_amount_input = QLineEdit()
        info_layout.addWidget(self.initial_amount_input)

        info_group.setLayout(info_layout)
        layout.addWidget(info_group)

        # 交易日输入
        date_group = QWidget()
        date_layout = QHBoxLayout()

        date_layout.addWidget(QLabel("交易日:"))
        self.trade_date_input = QDateEdit()
        self.trade_date_input.setCalendarPopup(True)
        self.trade_date_input.setDate(QDate.currentDate())
        date_layout.addWidget(self.trade_date_input)

        self.add_date_btn = QPushButton("添加交易日")
        self.add_date_btn.clicked.connect(self.add_trade_date)
        date_layout.addWidget(self.add_date_btn)

        self.batch_add_btn = QPushButton("批量添加交易日")
        self.batch_add_btn.clicked.connect(self.batch_add_trade_dates)
        date_layout.addWidget(self.batch_add_btn)

        self.clear_dates_btn = QPushButton("清空交易日")
        self.clear_dates_btn.clicked.connect(self.clear_trade_dates)
        date_layout.addWidget(self.clear_dates_btn)

        self.delete_date_btn = QPushButton("删除选中交易日")
        self.delete_date_btn.clicked.connect(self.delete_selected_trade_date)
        date_layout.addWidget(self.delete_date_btn)

        date_group.setLayout(date_layout)
        layout.addWidget(date_group)

        # 交易日列表显示
        self.trade_dates_table = QTableWidget()
        self.trade_dates_table.setColumnCount(1)
        self.trade_dates_table.setHorizontalHeaderLabels(["交易日"])
        self.trade_dates_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.trade_dates_table.itemChanged.connect(self.update_daily_reversal)
        layout.addWidget(self.trade_dates_table)

        # 操作按钮
        btn_group = QWidget()
        btn_layout = QHBoxLayout()

        self.save_btn = QPushButton("保存期权")
        self.save_btn.clicked.connect(self.save_option)
        btn_layout.addWidget(self.save_btn)

        self.update_btn = QPushButton("更新期权")
        self.update_btn.clicked.connect(self.update_option)
        btn_layout.addWidget(self.update_btn)

        self.clear_btn = QPushButton("清空输入")
        self.clear_btn.clicked.connect(self.clear_inputs)
        btn_layout.addWidget(self.clear_btn)

        self.delete_btn = QPushButton("删除期权")
        self.delete_btn.clicked.connect(self.delete_option)
        btn_layout.addWidget(self.delete_btn)

        btn_group.setLayout(btn_layout)
        layout.addWidget(btn_group)

        tab.setLayout(layout)

    def setup_query_tab(self, tab):
        layout = QVBoxLayout()

        query_group1 = QWidget()
        query_layout1 = QHBoxLayout()

        query_layout1.addWidget(QLabel("查询日期:"))
        self.query_date_input = QDateEdit()
        self.query_date_input.setCalendarPopup(True)
        self.query_date_input.setDate(QDate.currentDate())
        query_layout1.addWidget(self.query_date_input)

        query_layout1.addWidget(QLabel("期权名称:"))
        self.query_option_combo = QComboBox()
        self.query_option_combo.addItem("所有期权", None)
        query_layout1.addWidget(self.query_option_combo)

        self.query_btn = QPushButton("查询")
        self.query_btn.clicked.connect(self.normal_query)
        query_layout1.addWidget(self.query_btn)

        self.refresh_btn = QPushButton("重新获取数据")
        self.refresh_btn.clicked.connect(self.refresh_market_data)
        query_layout1.addWidget(self.refresh_btn)

        self.edit_close_price_btn = QPushButton("修改收盘价")
        self.edit_close_price_btn.clicked.connect(self.edit_close_price)
        query_layout1.addWidget(self.edit_close_price_btn)

        self.edit_close_amount_btn = QPushButton("修改平仓量")
        self.edit_close_amount_btn.clicked.connect(self.edit_close_amount)
        query_layout1.addWidget(self.edit_close_amount_btn)

        query_group1.setLayout(query_layout1)
        layout.addWidget(query_group1)

        keyword_group = QWidget()
        keyword_layout = QHBoxLayout()

        keyword_layout.addWidget(QLabel("关键词筛选:"))
        self.keyword_input = QLineEdit()
        self.keyword_input.setPlaceholderText("输入期权名称中的关键词...")
        keyword_layout.addWidget(self.keyword_input)

        self.keyword_query_btn = QPushButton("关键词查询")
        self.keyword_query_btn.clicked.connect(self.keyword_query)
        keyword_layout.addWidget(self.keyword_query_btn)

        keyword_layout.addStretch()  # 右侧留白，使控件左对齐
        keyword_group.setLayout(keyword_layout)
        layout.addWidget(keyword_group)

        # 查询结果表格 - 单个期权显示
        self.single_option_table = QTableWidget()
        self.single_option_table.setColumnCount(8)
        self.single_option_table.setHorizontalHeaderLabels([
            "日期", "期权名称", "执行价格", "每日冲回量", "收盘价",
            "实际成交量", "平仓量", "最新头寸"
        ])
        # 调整列宽：加宽期权名称列（索引1），其余列均分
        self.adjust_table_column_widths(self.single_option_table)
        layout.addWidget(self.single_option_table)

        # 分隔线
        self.separator = QFrame()
        self.separator.setFrameShape(QFrame.HLine)
        self.separator.setFrameShadow(QFrame.Sunken)
        layout.addWidget(self.separator)
        self.separator.hide()

        # 查询结果表格 - 所有期权显示
        self.active_options_label = QLabel("未到期期权:")
        layout.addWidget(self.active_options_label)
        self.active_options_label.hide()

        self.active_options_table = QTableWidget()
        self.active_options_table.setColumnCount(8)
        self.active_options_table.setHorizontalHeaderLabels([
            "日期", "期权名称", "执行价格", "每日冲回量", "收盘价",
            "实际成交量", "平仓量", "最新头寸"
        ])
        # 调整列宽
        self.adjust_table_column_widths(self.active_options_table)
        layout.addWidget(self.active_options_table)
        self.active_options_table.hide()

        # 分隔线
        self.separator2 = QFrame()
        self.separator2.setFrameShape(QFrame.HLine)
        self.separator2.setFrameShadow(QFrame.Sunken)
        layout.addWidget(self.separator2)
        self.separator2.hide()

        # 已到期期权表格
        self.expired_options_label = QLabel("已到期期权:")
        layout.addWidget(self.expired_options_label)
        self.expired_options_label.hide()

        self.expired_options_table = QTableWidget()
        self.expired_options_table.setColumnCount(8)
        self.expired_options_table.setHorizontalHeaderLabels([
            "到期日期", "期权名称", "执行价格", "每日冲回量", "收盘价",
            "实际成交量", "平仓量", "最新头寸"
        ])
        # 调整列宽
        self.adjust_table_column_widths(self.expired_options_table)
        layout.addWidget(self.expired_options_table)
        self.expired_options_table.hide()

        tab.setLayout(layout)

    def adjust_table_column_widths(self, table):
        """加宽期权名称列（索引1），其余列均分"""
        # 设置所有列默认模式为Stretch
        for col in range(table.columnCount()):
            table.horizontalHeader().setSectionResizeMode(col, QHeaderView.Stretch)

        table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Interactive)
        # 设置初始宽度为其他列的1.8倍
        table.setColumnWidth(1, 270)

    def setup_close_tab(self, tab):
        layout = QVBoxLayout()

        # 选择期权和日期
        select_group = QWidget()
        select_layout = QHBoxLayout()

        select_layout.addWidget(QLabel("期权名称:"))
        self.close_option_combo = QComboBox()
        self.close_option_combo.currentIndexChanged.connect(self.update_close_dates)
        select_layout.addWidget(self.close_option_combo)

        select_layout.addWidget(QLabel("平仓日期:"))
        self.close_date_combo = QComboBox()
        select_layout.addWidget(self.close_date_combo)

        select_group.setLayout(select_layout)
        layout.addWidget(select_group)

        # 平仓量输入
        amount_group = QWidget()
        amount_layout = QHBoxLayout()

        amount_layout.addWidget(QLabel("平仓量:"))
        self.close_amount_input = QLineEdit()
        amount_layout.addWidget(self.close_amount_input)

        amount_group.setLayout(amount_layout)
        layout.addWidget(amount_group)

        # 操作按钮
        self.close_btn = QPushButton("记录平仓")
        self.close_btn.clicked.connect(self.record_close)
        layout.addWidget(self.close_btn)

        tab.setLayout(layout)

    def batch_add_trade_dates(self):
        dialog = BatchAddDatesDialog(self)
        if dialog.exec_() == QDialog.Accepted:
            dates = dialog.get_dates()
            for date in dates:
                # 检查是否已存在
                exists = False
                for row in range(self.trade_dates_table.rowCount()):
                    item = self.trade_dates_table.item(row, 0)
                    if item and item.text() == date:
                        exists = True
                        break

                if not exists:
                    row = self.trade_dates_table.rowCount()
                    self.trade_dates_table.insertRow(row)
                    item = QTableWidgetItem(date)
                    item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                    self.trade_dates_table.setItem(row, 0, item)

            # 更新每日冲回量
            self.update_daily_reversal()

    def add_trade_date(self):
        date = self.trade_date_input.date().toString("yyyy-MM-dd")

        # 检查是否已存在
        for row in range(self.trade_dates_table.rowCount()):
            item = self.trade_dates_table.item(row, 0)
            if item and item.text() == date:
                QMessageBox.warning(self, "警告", "该交易日已添加!")
                return

        # 添加到表格
        row = self.trade_dates_table.rowCount()
        self.trade_dates_table.insertRow(row)
        item = QTableWidgetItem(date)
        item.setFlags(item.flags() & ~Qt.ItemIsEditable)
        self.trade_dates_table.setItem(row, 0, item)

    def clear_trade_dates(self):
        self.trade_dates_table.setRowCount(0)

    def delete_selected_trade_date(self):
        selected_items = self.trade_dates_table.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, "警告", "请先选择要删除的交易日!")
            return

        rows_to_delete = set()
        for item in selected_items:
            rows_to_delete.add(item.row())

        for row in sorted(rows_to_delete, reverse=True):
            self.trade_dates_table.removeRow(row)

    def update_daily_reversal(self):
        if not self.initial_amount_input.text():
            return

        try:
            initial_amount = float(self.initial_amount_input.text())
        except ValueError:
            return

        row_count = self.trade_dates_table.rowCount()
        if row_count > 0:
            daily_reversal = -initial_amount / row_count
            option_name = self.option_name_input.text().strip()
            if option_name in self.options:
                self.options[option_name]["daily_reversal"] = daily_reversal
                self.recalculate_option_from_date(self.options[option_name],
                                                  self.options[option_name]["trade_dates"][0])

    def save_option(self):
        name = self.option_name_input.text().strip()
        code = self.option_code_input.text().strip()
        strike_price = self.strike_price_input.text().strip()
        initial_amount = self.initial_amount_input.text().strip()

        if not name or not code or not strike_price or not initial_amount:
            QMessageBox.warning(self, "警告", "请填写所有必填字段!")
            return

        if self.trade_dates_table.rowCount() == 0:
            QMessageBox.warning(self, "警告", "请至少添加一个交易日!")
            return

        try:
            strike_price = float(strike_price)
            initial_amount = float(initial_amount)
        except ValueError:
            QMessageBox.warning(self, "警告", "执行价格和初始计提量必须是数字!")
            return

        trade_dates = []
        for row in range(self.trade_dates_table.rowCount()):
            item = self.trade_dates_table.item(row, 0)
            if item:
                trade_dates.append(item.text())

        if not trade_dates:
            QMessageBox.warning(self, "警告", "没有有效的交易日!")
            return

        daily_reversal = -initial_amount / len(trade_dates)

        option_data = {
            "name": name,
            "code": code,
            "strike_price": strike_price,
            "initial_amount": initial_amount,
            "trade_dates": trade_dates,
            "daily_reversal": daily_reversal,
            "close_prices": {},
            "actual_volumes": {},
            "close_amounts": {},
            "position_changes": {},
            "positions": {}
        }

        # 检查是否已有同名期权，如存在则询问是否覆盖
        if name in self.options:
            reply = QMessageBox.question(
                self, '确认覆盖',
                f'期权 "{name}" 已存在，是否覆盖?',
                QMessageBox.Yes | QMessageBox.No, QMessageBox.No
            )
            if reply == QMessageBox.No:
                return

        self.options[name] = option_data
        self.update_option_combos()
        QMessageBox.information(self, "成功", f"期权 {name} 已保存!")
        self.clear_inputs()
        self.save_data()

    def update_option(self):
        name = self.option_name_input.text().strip()
        if not name or name not in self.options:
            QMessageBox.warning(self, "警告", "请选择要修改的期权!")
            return

        code = self.option_code_input.text().strip()
        strike_price = self.strike_price_input.text().strip()
        initial_amount = self.initial_amount_input.text().strip()

        if not name or not code or not strike_price or not initial_amount:
            QMessageBox.warning(self, "警告", "请填写所有必填字段!")
            return

        if self.trade_dates_table.rowCount() == 0:
            QMessageBox.warning(self, "警告", "请至少添加一个交易日!")
            return

        try:
            strike_price = float(strike_price)
            initial_amount = float(initial_amount)
        except ValueError:
            QMessageBox.warning(self, "警告", "执行价格和初始计提量必须是数字!")
            return

        trade_dates = []
        for row in range(self.trade_dates_table.rowCount()):
            item = self.trade_dates_table.item(row, 0)
            if item:
                trade_dates.append(item.text())

        if not trade_dates:
            QMessageBox.warning(self, "警告", "没有有效的交易日!")
            return

        daily_reversal = -initial_amount / len(trade_dates)

        self.options[name]["code"] = code
        self.options[name]["strike_price"] = strike_price
        self.options[name]["initial_amount"] = initial_amount
        self.options[name]["trade_dates"] = trade_dates
        self.options[name]["daily_reversal"] = daily_reversal

        self.recalculate_option_from_date(self.options[name], trade_dates[0])
        QMessageBox.information(self, "成功", f"期权 {name} 已更新!")
        self.save_data()

    def delete_option(self):
        name = self.option_name_input.text().strip()
        if not name or name not in self.options:
            QMessageBox.warning(self, "警告", "请选择要删除的期权!")
            return

        reply = QMessageBox.question(self, '确认',
                                     f'确定要删除期权 {name} 吗?',
                                     QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if reply == QMessageBox.No:
            return

        del self.options[name]
        self.update_option_combos()
        self.clear_inputs()
        QMessageBox.information(self, "成功", f"期权 {name} 已删除!")
        self.save_data()

    def load_option_for_edit(self):
        option_name = self.option_select_combo.currentData()

        # 如果是"新建期权"，清空输入栏
        if not option_name:
            self.clear_inputs()
            return

        if option_name not in self.options:
            return

        option = self.options[option_name]

        self.option_name_input.setText(option["name"])
        self.option_code_input.setText(option["code"])
        self.strike_price_input.setText(str(option["strike_price"]))
        self.initial_amount_input.setText(str(option["initial_amount"]))

        self.trade_dates_table.setRowCount(0)
        for date in option["trade_dates"]:
            row = self.trade_dates_table.rowCount()
            self.trade_dates_table.insertRow(row)
            item = QTableWidgetItem(date)
            item.setFlags(item.flags() & ~Qt.ItemIsEditable)
            self.trade_dates_table.setItem(row, 0, item)

    def clear_inputs(self):
        self.option_name_input.clear()
        self.option_code_input.clear()
        self.strike_price_input.clear()
        self.initial_amount_input.clear()
        self.trade_dates_table.setRowCount(0)

    def update_option_combos(self):
        self.query_option_combo.clear()
        self.query_option_combo.addItem("所有期权", None)

        self.close_option_combo.clear()

        self.option_select_combo.clear()
        self.option_select_combo.addItem("新建期权", "")

        for name in sorted(self.options.keys()):
            self.query_option_combo.addItem(name, name)
            self.close_option_combo.addItem(name, name)
            self.option_select_combo.addItem(name, name)

    def update_close_dates(self):
        self.close_date_combo.clear()

        option_name = self.close_option_combo.currentData()
        if not option_name or option_name not in self.options:
            return

        option = self.options[option_name]
        for date in option["trade_dates"]:
            self.close_date_combo.addItem(date, date)

    def normal_query(self):
        """普通查询（按期权名称）"""
        if self.query_thread and self.query_thread.isRunning():
            QMessageBox.information(self, "提示", "查询正在进行中，请稍候...")
            return

        query_date = self.query_date_input.date().toString("yyyy-MM-dd")
        option_name = self.query_option_combo.currentData()

        # 初始化进度显示
        self.query_in_progress = True
        self.progress_bar.show()
        self.progress_bar.setValue(0)
        self.cancel_btn.show()
        self.disable_buttons_during_operation(True)

        # 创建并启动查询线程（非关键词查询）
        self.query_thread = QueryThread(self, query_date, option_name, is_keyword_query=False)

        # 连接信号和槽
        self.query_thread.progress_updated.connect(self.update_progress)
        self.query_thread.finished.connect(self.on_query_finished)
        self.query_thread.result_ready.connect(self.display_query_results)

        self.query_thread.start()

    def keyword_query(self):
        """关键词查询"""
        if self.query_thread and self.query_thread.isRunning():
            QMessageBox.information(self, "提示", "查询正在进行中，请稍候...")
            return

        query_date = self.query_date_input.date().toString("yyyy-MM-dd")
        keyword = self.keyword_input.text().strip()

        if not keyword:
            QMessageBox.warning(self, "警告", "请输入关键词后再查询!")
            return

        # 初始化进度显示
        self.query_in_progress = True
        self.progress_bar.show()
        self.progress_bar.setValue(0)
        self.cancel_btn.show()
        self.disable_buttons_during_operation(True)

        # 创建并启动查询线程（关键词查询）
        self.query_thread = QueryThread(self, query_date, keyword=keyword, is_keyword_query=True)

        # 连接信号和槽
        self.query_thread.progress_updated.connect(self.update_progress)
        self.query_thread.finished.connect(self.on_query_finished)
        self.query_thread.result_ready.connect(self.display_query_results)

        self.query_thread.start()

    def display_query_results(self, results):
        """显示查询结果到表格中"""
        # 清空所有表格
        self.single_option_table.setRowCount(0)
        self.active_options_table.setRowCount(0)
        self.expired_options_table.setRowCount(0)

        # 根据查询类型显示不同的表格
        if results["single_option"]:
            # 显示单个期权数据
            self.single_option_table.show()
            self.separator.hide()
            self.active_options_table.hide()
            self.active_options_label.hide()
            self.separator2.hide()
            self.expired_options_table.hide()
            self.expired_options_label.hide()

            for item in results["single_option"]:
                self.add_query_result_row(self.single_option_table, item["date"], item["option"], False)
        else:
            # 显示所有期权数据（分为未到期和已到期）
            self.single_option_table.hide()
            self.separator.show()

            if results["active_options"]:
                self.active_options_table.show()
                self.active_options_label.show()
                self.active_options_label.setText(f"未到期期权 ({results['active_count']}个):")

                for item in results["active_options"]:
                    self.add_query_result_row(self.active_options_table, item["date"], item["option"], True)
            else:
                self.active_options_table.hide()
                self.active_options_label.hide()

            self.separator2.show()

            if results["expired_options"]:
                self.expired_options_table.show()
                self.expired_options_label.show()
                self.expired_options_label.setText(f"已到期期权 ({results['expired_count']}个):")

                for item in results["expired_options"]:
                    self.add_query_result_row(self.expired_options_table, item["date"], item["option"], True)
            else:
                self.expired_options_table.hide()
                self.expired_options_label.hide()

    @pyqtSlot(int, str)
    def update_progress(self, value, message):
        """更新进度条和状态信息"""
        self.progress_bar.setValue(value)
        self.progress_label.setText(f"{message}")

    def on_query_finished(self, success, error_messages):
        """查询完成后的处理"""
        self.progress_bar.setValue(100 if success else 0)
        self.cancel_btn.hide()
        self.disable_buttons_during_operation(False)
        self.query_in_progress = False

        if success:
            self.progress_label.setText("查询完成")

            # 显示错误信息（如果有）
            if error_messages and not ("error" in error_messages):
                message_text = "以下期权存在无法获取的收盘价数据：\n\n"
                for option_name, dates in error_messages.items():
                    option = self.options.get(option_name)
                    display_name = option["name"] if option else option_name
                    dates_str = ", ".join(dates)
                    message_text += f"{display_name}：\n{dates_str}\n\n"

                QMessageBox.warning(self, "数据获取失败", message_text)
        else:
            error_msg = error_messages.get("error", "查询失败")
            self.progress_label.setText(error_msg)
            QMessageBox.warning(self, "查询失败", error_msg)

        self.query_thread = None

    def add_query_result_row(self, table, date, option, is_all_options_mode):
        row = table.rowCount()
        table.insertRow(row)

        # 日期列
        table.setItem(row, 0, QTableWidgetItem(date))

        # 期权名称
        table.setItem(row, 1, QTableWidgetItem(option["name"]))

        # 执行价格
        table.setItem(row, 2, QTableWidgetItem(f"{option['strike_price']:.2f}"))

        # 每日冲回量
        table.setItem(row, 3, QTableWidgetItem(f"{option['daily_reversal']:.2f}"))

        # 收盘价
        close_price = option["close_prices"].get(date, "N/A")
        close_price_text = f"{close_price:.2f}" if close_price != "N/A" else "N/A"
        table.setItem(row, 4, QTableWidgetItem(close_price_text))

        # 实际成交量
        actual_volume = option["actual_volumes"].get(date, 0)
        table.setItem(row, 5, QTableWidgetItem(f"{actual_volume:.2f}"))

        # 平仓量
        close_amount = option["close_amounts"].get(date, 0)
        table.setItem(row, 6, QTableWidgetItem(f"{close_amount:.2f}"))

        # 最新头寸
        position = option["positions"].get(date, 0)
        table.setItem(row, 7, QTableWidgetItem(f"{position:.2f}"))

    def calculate_option_data(self, option, end_date=None):
        """计算期权数据，只计算到指定日期"""
        prev_position = option["initial_amount"]

        for date in option["trade_dates"]:
            if end_date and date > end_date:
                break

            if date not in option["close_prices"]:
                # 只获取查询日期及之前的数据
                if date <= end_date if end_date else True:
                    close_price = self.get_dce_daily_close(option["code"], date)
                    option["close_prices"][date] = close_price if close_price is not None else "N/A"

            if date not in option["actual_volumes"]:
                actual_volume = 0
                close_price = option["close_prices"].get(date, "N/A")

                if close_price != "N/A":
                    if option["initial_amount"] < 0:
                        if close_price > option["strike_price"]:
                            actual_volume = -option["daily_reversal"]
                    else:
                        if close_price < option["strike_price"]:
                            actual_volume = -option["daily_reversal"]

                option["actual_volumes"][date] = actual_volume

            close_amount = option["close_amounts"].get(date, 0)

            position_change = option["daily_reversal"] + option["actual_volumes"][date]
            option["position_changes"][date] = position_change

            current_position = prev_position + position_change + close_amount
            option["positions"][date] = current_position
            prev_position = current_position

    def edit_close_price(self):
        # 检查哪个表格有选中项
        selected_items = self.single_option_table.selectedItems()
        table = self.single_option_table
        if not selected_items or len(selected_items) < 5:
            selected_items = self.active_options_table.selectedItems()
            table = self.active_options_table
            if not selected_items or len(selected_items) < 5:
                selected_items = self.expired_options_table.selectedItems()
                table = self.expired_options_table
                if not selected_items or len(selected_items) < 5:
                    QMessageBox.warning(self, "警告", "请先选择要修改的行!")
                    return

        row = selected_items[0].row()
        option_name = table.item(row, 1).text()
        date = table.item(row, 0).text()

        if option_name not in self.options:
            QMessageBox.warning(self, "警告", "找不到对应的期权数据!")
            return

        current_price_item = table.item(row, 4)
        current_price = current_price_item.text()

        try:
            if current_price == "N/A":
                default_value = 0.0
            else:
                default_value = float(current_price)
        except ValueError:
            default_value = 0.0

        new_price, ok = QInputDialog.getDouble(
            self, "修改收盘价",
            f"请输入 {option_name} 在 {date} 的新收盘价:",
            value=default_value,
            min=0.0
        )

        if ok:
            self.options[option_name]["close_prices"][date] = new_price
            self.recalculate_option_from_date(self.options[option_name], date)
            # 保存修改后的数据
            self.save_data()
            # 根据当前查询类型重新查询
            self.requery_based_on_last_action()
            QMessageBox.information(self, "成功", "收盘价已更新并重新计算!")

    def edit_close_amount(self):
        # 检查哪个表格有选中项
        selected_items = self.single_option_table.selectedItems()
        table = self.single_option_table
        if not selected_items or len(selected_items) < 7:
            selected_items = self.active_options_table.selectedItems()
            table = self.active_options_table
            if not selected_items or len(selected_items) < 7:
                selected_items = self.expired_options_table.selectedItems()
                table = self.expired_options_table
                if not selected_items or len(selected_items) < 7:
                    QMessageBox.warning(self, "警告", "请先选择要修改的行!")
                    return

        row = selected_items[0].row()
        option_name = table.item(row, 1).text()
        date = table.item(row, 0).text()

        if option_name not in self.options:
            QMessageBox.warning(self, "警告", "找不到对应的期权数据!")
            return

        current_amount_item = table.item(row, 6)
        current_amount = current_amount_item.text()

        try:
            default_value = float(current_amount)
        except ValueError:
            default_value = 0.0

        new_amount, ok = QInputDialog.getDouble(
            self, "修改平仓量",
            f"请输入 {option_name} 在 {date} 的新平仓量:",
            value=default_value
        )

        if ok:
            self.options[option_name]["close_amounts"][date] = new_amount
            self.recalculate_option_from_date(self.options[option_name], date)
            # 保存修改后的数据
            self.save_data()
            # 根据当前查询类型重新查询
            self.requery_based_on_last_action()
            QMessageBox.information(self, "成功", "平仓量已更新并重新计算!")

    def refresh_market_data(self):
        """重新从网站获取数据并更新，使用线程避免UI卡顿"""
        if self.refresh_thread and self.refresh_thread.isRunning():
            QMessageBox.information(self, "提示", "数据更新正在进行中，请稍候...")
            return

        query_date = self.query_date_input.date().toString("yyyy-MM-dd")
        option_name = self.query_option_combo.currentData()
        keyword = self.keyword_input.text().strip()

        # 初始化进度显示
        self.progress_bar.show()
        self.progress_bar.setValue(0)
        self.cancel_btn.show()
        self.disable_buttons_during_operation(True)

        # 创建并启动线程，传入关键词以便筛选需要刷新的数据
        self.refresh_thread = DataRefreshThread(
            self,
            option_name if not keyword else None,
            query_date,
            keyword  # 传递关键词用于筛选
        )

        # 连接信号和槽
        self.refresh_thread.progress_updated.connect(self.update_progress)
        self.refresh_thread.finished.connect(self.on_refresh_finished)

        self.refresh_thread.start()

    def on_refresh_finished(self, success, message):
        """刷新完成后的处理"""
        self.progress_bar.setValue(100 if success else 0)
        self.progress_label.setText(message)
        self.cancel_btn.hide()
        self.disable_buttons_during_operation(False)

        if success:
            # 根据最后一次查询类型重新查询
            self.requery_based_on_last_action()

        self.refresh_thread = None

    def requery_based_on_last_action(self):
        """根据最后一次查询类型重新查询"""
        # 判断最后一次是普通查询还是关键词查询
        if self.query_thread and hasattr(self.query_thread, 'is_keyword_query'):
            if self.query_thread.is_keyword_query:
                # 重新执行关键词查询
                self.keyword_query()
            else:
                # 重新执行普通查询
                self.normal_query()
        else:
            # 默认执行普通查询
            self.normal_query()

    def cancel_operation(self):
        """取消当前正在进行的操作"""
        if self.refresh_thread and self.refresh_thread.isRunning():
            self.refresh_thread.cancel()
            self.progress_label.setText("正在取消操作...")
        elif self.query_thread and self.query_thread.isRunning():
            self.query_thread.cancel()
            self.progress_label.setText("正在取消查询...")
        elif self.query_in_progress:
            self.query_in_progress = False
            self.progress_label.setText("查询已取消")
            self.cancel_btn.hide()
            self.disable_buttons_during_operation(False)

    def disable_buttons_during_operation(self, disable):
        """在操作进行时禁用相关按钮"""
        # 查询标签页按钮
        self.query_btn.setEnabled(not disable)
        self.keyword_query_btn.setEnabled(not disable)
        self.refresh_btn.setEnabled(not disable)
        self.edit_close_price_btn.setEnabled(not disable)
        self.edit_close_amount_btn.setEnabled(not disable)

        # 其他标签页按钮
        self.save_btn.setEnabled(not disable)
        self.update_btn.setEnabled(not disable)
        self.delete_btn.setEnabled(not disable)
        self.add_date_btn.setEnabled(not disable)
        self.batch_add_btn.setEnabled(not disable)
        self.clear_dates_btn.setEnabled(not disable)
        self.delete_date_btn.setEnabled(not disable)
        self.close_btn.setEnabled(not disable)

    def refresh_option_data(self, option, date):
        """刷新单个期权的市场数据"""
        # 重新获取收盘价
        close_price = self.get_dce_daily_close(option["code"], date)
        if close_price is not None:
            option["close_prices"][date] = close_price
            # 重新计算该日期及之后的数据
            self.recalculate_option_from_date(option, date)
        else:
            # 查询日期前的数据获取失败才标记为N/A
            current_date = QDate.currentDate().toString("yyyy-MM-dd")
            if date < current_date:
                option["close_prices"][date] = "N/A"
            self.recalculate_option_from_date(option, date)

    def record_close(self):
        option_name = self.close_option_combo.currentData()
        if not option_name or option_name not in self.options:
            QMessageBox.warning(self, "警告", "请选择有效的期权!")
            return

        date = self.close_date_combo.currentData()
        if not date:
            QMessageBox.warning(self, "警告", "请选择有效的日期!")
            return

        try:
            close_amount = float(self.close_amount_input.text())
        except ValueError:
            QMessageBox.warning(self, "警告", "请输入有效的平仓量!")
            return

        option = self.options[option_name]
        option["close_amounts"][date] = close_amount
        self.recalculate_option_from_date(option, date)

        QMessageBox.information(self, "成功", f"已记录 {option_name} 在 {date} 的平仓量 {close_amount}")
        self.close_amount_input.clear()
        self.save_data()

    def recalculate_option_from_date(self, option, start_date):
        if start_date not in option["trade_dates"]:
            return

        start_index = option["trade_dates"].index(start_date)

        prev_position = option["initial_amount"]
        if start_index > 0:
            prev_date = option["trade_dates"][start_index - 1]
            prev_position = option["positions"].get(prev_date, option["initial_amount"])

        for i in range(start_index, len(option["trade_dates"])):
            date = option["trade_dates"][i]

            actual_volume = 0
            close_price = option["close_prices"].get(date, "N/A")

            if close_price != "N/A":
                if option["initial_amount"] < 0:
                    if close_price > option["strike_price"]:
                        actual_volume = -option["daily_reversal"]
                else:
                    if close_price < option["strike_price"]:
                        actual_volume = -option["daily_reversal"]

            option["actual_volumes"][date] = actual_volume

            close_amount = option["close_amounts"].get(date, 0)

            position_change = option["daily_reversal"] + option["actual_volumes"][date]
            option["position_changes"][date] = position_change

            current_position = prev_position + position_change + close_amount
            option["positions"][date] = current_position
            prev_position = current_position

    def save_data(self):
        try:
            data_to_save = {}
            for name, option in self.options.items():
                data_to_save[name] = {
                    "name": option["name"],
                    "code": option["code"],
                    "strike_price": option["strike_price"],
                    "initial_amount": option["initial_amount"],
                    "trade_dates": option["trade_dates"],
                    "daily_reversal": option["daily_reversal"],
                    "close_prices": option["close_prices"],
                    "actual_volumes": option["actual_volumes"],
                    "close_amounts": option["close_amounts"],
                    "position_changes": option["position_changes"],
                    "positions": option["positions"]
                }

            with open(self.data_file, 'w') as f:
                json.dump(data_to_save, f, indent=4)

            return True
        except Exception as e:
            QMessageBox.warning(self, "错误", f"保存数据失败: {str(e)}")
            return False

    def save_data_as(self):
        file_name, _ = QFileDialog.getSaveFileName(self, "保存数据", "", "JSON文件 (*.json)")
        if file_name:
            self.data_file = file_name
            if self.save_data():
                QMessageBox.information(self, "成功", f"数据已保存到 {file_name}")

    def load_data(self):
        try:
            with open(self.data_file, 'r') as f:
                data = json.load(f)

            self.options = {}
            for name, option_data in data.items():
                self.options[name] = {
                    "name": option_data["name"],
                    "code": option_data["code"],
                    "strike_price": option_data["strike_price"],
                    "initial_amount": option_data["initial_amount"],
                    "trade_dates": option_data["trade_dates"],
                    "daily_reversal": option_data["daily_reversal"],
                    "close_prices": option_data["close_prices"],
                    "actual_volumes": option_data["actual_volumes"],
                    "close_amounts": option_data["close_amounts"],
                    "position_changes": option_data["position_changes"],
                    "positions": option_data["positions"]
                }

            self.update_option_combos()
            return True
        except FileNotFoundError:
            return False
        except Exception as e:
            QMessageBox.warning(self, "错误", f"加载数据失败: {str(e)}")
            return False

    def load_data_from_file(self):
        file_name, _ = QFileDialog.getOpenFileName(self, "加载数据", "", "JSON文件 (*.json)")
        if file_name:
            self.data_file = file_name
            if self.load_data():
                QMessageBox.information(self, "成功", f"已从 {file_name} 加载数据")
            else:
                QMessageBox.warning(self, "错误", "加载数据失败")

    def get_dce_daily_close(self, contract_code: str, date_yyyymmdd: str) -> float | None:
        date_yyyymmdd = date_yyyymmdd.replace("-", "")

        url = "http://www.dce.com.cn/publicweb/quotesdata/dayQuotesCh.html"
        params = {
            "dayQuotes.variety": "all",
            "dayQuotes.trade_type": "0",
            "year": date_yyyymmdd[:4],
            "month": str(int(date_yyyymmdd[4:6]) - 1),
            "day": date_yyyymmdd[6:8],
        }

        try:
            response = requests.get(
                url,
                params=params,
                headers={"User-Agent": "Mozilla/5.0"},
                timeout=10
            )
            response.encoding = 'utf-8'

            if "大连商品交易所  日行情表" not in response.text:
                return None

            df = pd.read_html(StringIO(response.text), header=0)[0]
            df.columns = [col.strip() for col in df.columns]

            if '合约名称' not in df.columns or '收盘价' not in df.columns:
                return None

            df['合约名称'] = df['合约名称'].astype(str).str.strip()
            target_row = df[df['合约名称'].str.lower() == contract_code.strip().lower()]

            if target_row.empty:
                return None

            close_price = target_row.iloc[0]['收盘价']
            return float(close_price) if pd.notna(close_price) and close_price != "-" else None

        except Exception:
            return None


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = OptionPositionCalculator()
    window.show()
    sys.exit(app.exec_())
