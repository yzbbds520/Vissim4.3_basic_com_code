# Python脚本用于设置Vissim 4.3/4.5中的车辆输入流量
import win32com.client as com

try:
    # 连接到VISSIM
    print("正在连接到VISSIM...")
    Vissim = com.Dispatch("VISSIM.Vissim")
    print("成功连接到VISSIM")

    # 加载网络文件
    Filename = r"D:\python\project\vissim\try.inp"
    print(f"尝试加载网络文件: {Filename}")
    Vissim.LoadNet(Filename)
    print("成功加载网络文件")

    # 获取VehicleInputs
    vi_count = Vissim.Net.VehicleInputs.Count
    print(f"当前VehicleInputs数量: {vi_count}")

    # 获取第一个车辆输入
    if vi_count > 0:
        first_vi = Vissim.Net.VehicleInputs.Item(1)

        # 获取当前流量
        current_volume = first_vi.AttValue('Volume')
        print(f"当前流量值: {current_volume}")

        # 设置新流量
        new_volume = 1000
        first_vi.SetAttValue('Volume', new_volume)
        print(f"设置流量为: {new_volume}")

        # 验证流量是否已更改
        updated_volume = first_vi.AttValue('Volume')
        print(f"更新后流量值: {updated_volume}")

        print("开始运行仿真...")
        Vissim.Simulation.RunContinuous()
        print("仿真已完成")

    else:
        print("没有车辆输入。您可能需要先添加车辆输入。")



except Exception as e:
    print(f"发生错误: {str(e)}")

finally:
    try:
        # 询问是否保存
        save = input("\n是否保存网络? (y/n): ")
        if save.lower() == 'y':
            save_path = r"D:\python\project\vissim\try_modified.inp"
            Vissim.SaveNet(save_path)
            print(f"网络已保存至: {save_path}")



        # 关闭VISSIM
        print("正在关闭VISSIM...")
        Vissim = None
        print("VISSIM已关闭")
    except:
        print("关闭VISSIM时出错")