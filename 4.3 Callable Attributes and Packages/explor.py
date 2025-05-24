# Python脚本用于探索Vissim 4.3/4.5的COM接口
import win32com.client as com
import pythoncom
import sys


# 辅助函数：打印对象的方法和属性
def explore_object(obj, obj_name):
    print(f"\n{'=' * 50}")
    print(f"探索对象: {obj_name}")
    print(f"{'=' * 50}")

    # 尝试获取对象的所有属性和方法
    try:
        attrs = dir(obj)
        print(f"\n对象 {obj_name} 的属性和方法:")
        for attr in attrs:
            if not attr.startswith('_'):  # 忽略内部属性
                print(f"  - {attr}")
    except:
        print(f"无法获取 {obj_name} 的属性和方法列表")

    # 尝试获取对象的类型信息
    try:
        type_info = str(type(obj))
        print(f"\n对象类型: {type_info}")
    except:
        print("无法获取对象类型信息")


try:
    # 连接到VISSIM
    print("正在连接到VISSIM 4.3/4.5...")
    Vissim = com.Dispatch("VISSIM.Vissim")
    print("成功连接到VISSIM")

    # 加载网络文件
    try:
        Filename = r"D:\python\project\vissim\try.inp"
        print(f"尝试加载网络文件: {Filename}")
        Vissim.LoadNet(Filename)
        print("成功加载网络文件")
    except Exception as e:
        print(f"加载网络文件失败: {str(e)}")

    # 探索主Vissim对象
    explore_object(Vissim, "Vissim")

    # 尝试探索Simulation对象
    try:
        if hasattr(Vissim, 'Simulation'):
            print("\n找到Simulation对象")
            explore_object(Vissim.Simulation, "Vissim.Simulation")
        else:
            print("\n未找到Simulation对象，尝试其他可能的属性名...")

            # 尝试其他可能的名称
            possible_names = ['Simulation', 'Sim', 'Simulator', 'SimulationRun']
            for name in possible_names:
                try:
                    if hasattr(Vissim, name):
                        print(f"找到可能的Simulation对象: {name}")
                        explore_object(getattr(Vissim, name), f"Vissim.{name}")
                except:
                    pass
    except Exception as e:
        print(f"探索Simulation对象时出错: {str(e)}")

    # 尝试找出如何设置随机种子
    try:
        print("\n尝试不同方式设置随机种子:")

        # 方法1：直接在Simulation对象上设置
        try:
            if hasattr(Vissim, 'Simulation'):
                print("尝试: Vissim.Simulation.RandSeed = 42")
                Vissim.Simulation.RandSeed = 42
                print("  - 成功!")
        except Exception as e:
            print(f"  - 失败: {str(e)}")

        # 方法2：使用SetAttValue
        try:
            if hasattr(Vissim, 'Simulation') and hasattr(Vissim.Simulation, 'SetAttValue'):
                print("尝试: Vissim.Simulation.SetAttValue('RandSeed', 42)")
                Vissim.Simulation.SetAttValue('RandSeed', 42)
                print("  - 成功!")
        except Exception as e:
            print(f"  - 失败: {str(e)}")

        # 方法3：在Net对象上设置
        try:
            if hasattr(Vissim, 'Net') and hasattr(Vissim.Net, 'SetAttValue'):
                print("尝试: Vissim.Net.SetAttValue('RandSeed', 42)")
                Vissim.Net.SetAttValue('RandSeed', 42)
                print("  - 成功!")
        except Exception as e:
            print(f"  - 失败: {str(e)}")

    except Exception as e:
        print(f"测试随机种子设置时出错: {str(e)}")

    # 尝试找出如何运行仿真
    try:
        print("\n尝试不同方式运行仿真:")

        # 方法1
        try:
            if hasattr(Vissim, 'Simulation') and hasattr(Vissim.Simulation, 'RunContinuous'):
                print("尝试: Vissim.Simulation.RunContinuous()")
                # 注释掉实际运行，以避免程序被阻塞
                # Vissim.Simulation.RunContinuous()
                print("  - 方法存在!")
        except Exception as e:
            print(f"  - 方法不存在: {str(e)}")

        # 方法2
        try:
            if hasattr(Vissim, 'Simulation') and hasattr(Vissim.Simulation, 'Run'):
                print("尝试: Vissim.Simulation.Run()")
                # 注释掉实际运行，以避免程序被阻塞
                # Vissim.Simulation.Run()
                print("  - 方法存在!")
        except Exception as e:
            print(f"  - 方法不存在: {str(e)}")

    except Exception as e:
        print(f"测试仿真运行方法时出错: {str(e)}")

except Exception as e:
    print(f"发生错误: {str(e)}")
finally:
    try:
        # 尝试关闭VISSIM
        print("\n尝试关闭VISSIM...")
        Vissim = None
        print("VISSIM已关闭")
    except:
        pass