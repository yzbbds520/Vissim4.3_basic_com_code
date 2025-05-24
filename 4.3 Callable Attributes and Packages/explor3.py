# Python脚本尝试在Vissim 4.3/4.5中创建链接
import win32com.client as com
import pythoncom

try:
    # 连接到VISSIM
    print("正在连接到VISSIM...")
    Vissim = com.Dispatch("VISSIM.Vissim")
    print("成功连接到VISSIM")

    # 创建新网络或加载现有网络
    try:
        Filename = r"D:\python\project\vissim\try.inp"
        print(f"尝试加载网络文件: {Filename}")
        Vissim.LoadNet(Filename)
        print("成功加载网络文件")
    except:
        print("尝试创建新网络...")
        Vissim.New()
        print("已创建新网络")

    # 尝试方法1：使用COM的低级Invoke方法
    print("\n方法1：尝试使用COM的低级Invoke方法")
    try:
        # 常见的方法名称可能是AddLink或CreateLink或Add
        method_name = "Add"  # 尝试不同的名称
        # 典型参数可能是：起始x,y和终止x,y坐标，车道数
        param1 = 0  # 起始x
        param2 = 0  # 起始y
        param3 = 0  # 终止x
        param4 = 50  # 终止y
        param5 = 1  # 车道数

        # 尝试调用未发现的方法
        DISPATCH_METHOD = 1
        result = Vissim.Net.Links._oleobj_.Invoke(
            Vissim.Net.Links._GetIDsOfNames_(method_name)[0],
            0, DISPATCH_METHOD, 1,
            param1, param2, param3, param4, param5
        )
        print(f"调用成功！结果: {result}")
    except Exception as e:
        print(f"方法1失败: {str(e)}")

    # 尝试方法2：使用Item方法并尝试创建
    print("\n方法2：检查现有链接")
    try:
        # 获取当前链接数量
        links_count = Vissim.Net.Links.Count
        print(f"当前链接数量: {links_count}")

        # 检查是否有链接
        if links_count > 0:
            # 获取第一个链接并查看其属性
            first_link = Vissim.Net.Links.Item(1)  # 或使用GetLinkByNumber
            print("获取第一个链接的属性...")
            link_attrs = dir(first_link)
            for attr in link_attrs:
                if not attr.startswith('_'):
                    print(f"  - {attr}")

            # 检查链接的坐标点
            if hasattr(first_link, 'Points'):
                print("\n链接有Points属性")
                points = first_link.Points
                print(f"链接的点数量: {points.Count if hasattr(points, 'Count') else '未知'}")
            elif hasattr(first_link, 'AttValue'):
                print("\n尝试通过AttValue获取链接信息")
                try:
                    # 尝试常见的属性名
                    possible_attrs = ['Start', 'End', 'StartX', 'StartY', 'EndX', 'EndY', 'Length']
                    for attr in possible_attrs:
                        try:
                            value = first_link.AttValue(attr)
                            print(f"  - {attr}: {value}")
                        except:
                            pass
                except Exception as e:
                    print(f"获取链接属性失败: {str(e)}")
    except Exception as e:
        print(f"方法2失败: {str(e)}")

    # 方法3：尝试通过可视化操作创建链接
    print("\n方法3：可能的解决方案")
    print("1. 使用Vissim GUI手动创建链接，然后保存网络")
    print("2. 查阅Vissim 4.3/4.5的COM编程手册获取正确的方法")
    print("3. 使用Vissim的导入/导出功能：将网络导出为文本格式，修改后再导入")

    # 获取可用的COM方法（深度探索）
    print("\n深度探索可用的COM方法:")
    try:
        print("检查Vissim.Net上可用的所有方法和属性...")
        net_attrs = dir(Vissim.Net)
        com_methods = [attr for attr in net_attrs if not attr.startswith('_')]
        print("COM方法和属性:")
        for method in sorted(com_methods):
            print(f"  - {method}")
    except Exception as e:
        print(f"深度探索失败: {str(e)}")

except Exception as e:
    print(f"发生错误: {str(e)}")
finally:
    try:
        # 询问是否保存
        save = input("\n是否保存网络? (y/n): ")
        if save.lower() == 'y':
            save_path = r"D:\python\project\vissim\try_modified.inp"
            Vissim.SaveNetAs(save_path)
            print(f"网络已保存至: {save_path}")

        # 关闭VISSIM
        print("正在关闭VISSIM...")
        Vissim = None
        print("VISSIM已关闭")
    except:
        print("关闭VISSIM时出错")