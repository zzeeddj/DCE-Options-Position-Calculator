# 期权头寸计算及期货数据统计系统

![Python](https://img.shields.io/badge/Python-3.7%2B-blue)
![PyQt5](https://img.shields.io/badge/GUI-PyQt5-green)
![License](https://img.shields.io/badge/License-MIT-orange)

## 目录
- [简介](#简介)
- [功能特点](#功能特点)
- [安装依赖](#安装依赖)
  - [安装方式](#安装方式)
  - [安装依赖](#安装依赖)
- [使用说明](#使用说明)
  - [期权管理](#期权管理)
  - [数据查询](#数据查询)
  - [平仓操作](#平仓操作)
  - [数据维护](#数据维护)
- [数据文件](#数据文件)
- [常见问题](#常见问题)
- [开发信息](#开发信息)
  
## 简介

本系统是一个用于计算期权头寸并统计期货数据的GUI应用程序，主要面向大连商品交易所(DCE)的期权产品。系统可以自动从DCE官网获取每日收盘价数据，计算期权的每日冲回量、实际成交量、平仓量和最新头寸等关键指标。

## 功能特点

- **期权管理**：添加、修改、删除期权合约，设置执行价格和初始计提量
- **交易日管理**：支持单个或批量添加交易日，自动计算每日冲回量
- **数据查询**：按日期查询期权数据，支持关键词筛选
- **平仓操作**：记录平仓量并自动重新计算头寸
- **数据刷新**：从DCE官网重新获取收盘价数据
- **数据持久化**：所有数据自动保存到JSON文件
- **多线程操作**：后台执行数据获取和计算，避免UI卡顿

## 系统要求

- Python 3.7+
- Windows/Linux/macOS
- PyQt5
- pandas
- requests

### 安装方式
克隆仓库或下载源代码
```bash
git clone https://github.com/your-repo/option-position-calculator.git
cd option-position-calculator


