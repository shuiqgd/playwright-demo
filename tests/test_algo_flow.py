from playwright.sync_api import Page, expect, sync_playwright
import time
import os

# 全局存储测试用例结果
test_results = []

def record_result(module, step, status, memo=""):
    """记录单独一条测试步骤的执行结果"""
    test_results.append({
        "module": module,
        "step": step,
        "status": status,
        "memo": memo
    })
    print(f"[{status}] {step}")

def generate_report():
    """测试结束后将结果写入轻量级 HTML 报告"""
    report_file = os.path.join(os.getcwd(), "Test_Report.html")
    
    html_content = f"""
    <!DOCTYPE html>
    <html lang="zh-CN">
    <head>
        <meta charset="utf-8">
        <title>AI编排器 UI 自动化测试报告</title>
        <style>
            body {{ font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; margin: 40px; background-color: #f7f9fc; }}
            h1 {{ color: #2c3e50; text-align: center; }}
            .summary {{ text-align: center; margin-bottom: 30px; color: #7f8c8d; }}
            table {{ width: 100%; border-collapse: collapse; background: #fff; box-shadow: 0 1px 3px rgba(0,0,0,0.1); }}
            th, td {{ border: 1px solid #e1e8ed; padding: 15px; text-align: left; }}
            th {{ background-color: #34495e; color: #fff; }}
            tr:nth-child(even) {{ background-color: #f9fbfc; }}
            .Pass {{ color: #27ae60; font-weight: bold; }}
            .Fail {{ color: #c0392b; font-weight: bold; }}
        </style>
    </head>
    <body>
        <h1>📊 UI 自动化测试执行报告</h1>
        <div class="summary">执行时间: {time.strftime('%Y-%m-%d %H:%M:%S')} &nbsp;&nbsp;|&nbsp;&nbsp; 目标系统: AI算法编排器系统</div>
        <table>
            <tr>
                <th>测试模块</th>
                <th>操作步骤与断言内容</th>
                <th>执行结果</th>
                <th>备注及细节</th>
            </tr>
    """
    
    for item in test_results:
        status_class = "Pass" if item['status'] == "Pass" else "Fail"
        result_text = "✅ 成功" if item['status'] == "Pass" else "❌ 失败"
        html_content += f"""
            <tr>
                <td>{item['module']}</td>
                <td>{item['step']}</td>
                <td class='{status_class}'>{result_text}</td>
                <td>{item['memo']}</td>
            </tr>
        """
        
    html_content += """
        </table>
    </body>
    </html>
    """
    
    with open(report_file, "w", encoding="utf-8") as f:
        f.write(html_content)
    print(f"\n======================================")
    print(f"👉 报告已成功生成并保存至: {report_file}")
    print(f"======================================")


def test_main_transformer_climbing_scenario(page: Page):
    """
    【功能】主变登高场景加载与验证 (接管本地已登录浏览器)
    """
    # ==========================
    # 模块 1: 登录状态及主页加载 (兼容掉线重登)
    # ==========================
    try:
        print("访问主页，检查是否保持登录态...")
        page.goto("http://192.168.88.98:12000/child/algoFlow/#/ai-orchestrator/workbench")
        
        try:
            # 先试探性地等5秒看进去了没
            expect(page.get_by_text("AI算法编排器").first).to_be_visible(timeout=5000)
            record_result("基础核心", "工作台页面加载与登录状态检测", "Pass", "本地Cookies仍有效，免密直接进入")
        except Exception:
            print("检测到本地未登录或登录已过期，正自动触发 SSO 转接登录流程...")
            # ================= [接回之前的超长容忍策略] =================
            print("→ [延时处理] 正给予页面充足时间进行 SSO 处理及重定向...")
            page.wait_for_timeout(5000) 
            
            print("→ 正在不限时等待最终落定的带有 input 输入框的登录页面出现...")
            page.wait_for_selector("input", timeout=0)

            print("→ 已到达最终的登录页面，输入账号: admin，密码: admin123")
            page.locator("input").first.fill("admin")
            page.locator("input[type='password']").first.fill("admin123")
            
            page.locator("button", has_text="登录").first.click()

            print("1. 正在等待系统授权并跳转回工作台页面...")
            expect(page.get_by_text("AI算法编排器").first).to_be_visible(timeout=20000)
            record_result("基础核心", "工作台页面加载与登录状态检测", "Pass", "已完成SSO重新授权并进入系统")
            
    except Exception as e:
        record_result("基础核心", "工作台页面加载与登录状态检测", "Fail", str(e))
        raise e

    # ==========================
    # 模块 2: 主变登高场景渲染断言
    # ==========================
    try:
        scene_button = page.get_by_text("主变登高场景", exact=True)
        expect(scene_button).to_be_visible(timeout=5000)
        scene_button.click()

        expect(page.get_by_text("配置面板")).to_be_visible(timeout=5000)
        
        # 验证四个核心算子节点是否在幕布渲染出
        start_node = page.get_by_text("开始", exact=True).first
        intrusion_node = page.get_by_text("区域入侵检测算子").first
        ticket_node = page.get_by_text("工作票识别算子").first
        alarm_node = page.get_by_text("诊断报警").first
        
        expect(start_node).to_be_visible(timeout=5000)
        expect(intrusion_node).to_be_visible(timeout=5000)
        expect(ticket_node).to_be_visible(timeout=5000)
        expect(alarm_node).to_be_visible(timeout=5000)
        record_result("流程可视化", "主变登高场景核心节点树加载与验证", "Pass", "四大流程关键节点(开始,入侵检测,工作票,报警)均正常呈现")
    except Exception as e:
        record_result("流程可视化", "主变登高场景核心节点树加载与验证", "Fail", str(e))
        raise e

    # ==========================
    # 模块 3: 多菜单切换遍历
    # ==========================
    menu_items = [
        "算子仓库", 
        "原始报警", 
        "诊断报警", 
        "IV-CV诊断报告", 
        "摄像头管理", 
        "操作日志",
        "AI算法编排器"  # 最后切回主视图
    ]
    
    for menu in menu_items:
        try:
            menu_locator = page.get_by_text(menu, exact=True).first
            expect(menu_locator).to_be_visible(timeout=5000)
            menu_locator.click()
            # 视觉过渡停顿，更方便肉眼校验
            page.wait_for_timeout(1500)
            record_result("功能遍历", f"成功切换至左侧菜单: {menu}", "Pass", "组件及路由加载正常")
        except Exception as e:
            record_result("功能遍历", f"切换侧边栏菜单 {menu} 时发生异常", "Fail", str(e))
            raise e


if __name__ == "__main__":
    print("🚀 开始进行包含报表生成的自动化测试连跑...")
    with sync_playwright() as p:
        try:
            # 绝不冷启动无头浏览器，而是连接本地 9222 调试端口复用 Cookie
            browser = p.chromium.connect_over_cdp("http://localhost:9222")
            context = browser.contexts[0]
            # 开辟一个新标签页来跑这次测试以防污染
            page = context.new_page()
            
            test_main_transformer_climbing_scenario(page)
            
        except Exception as e:
            print(f"❌ 测试发生致命错误: {e}")
        finally:
            print("🛑 指令执行退出。正在汇总并输出测试报告数据...")
            # 不论中间哪一步抛出了异常崩溃，最终一并生成测试报告
            generate_report()
            
            # 不关闭源浏览器 context.close()
            # 仅关闭我们为了本次测试新打开的标签页
            if 'page' in locals() and not page.is_closed():
                page.close()
