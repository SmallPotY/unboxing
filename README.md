# 箱型推荐

[TOC]

### 一、功能说明

```python
# 用户自行可维护商品的长宽高,定义箱型的品牌与长宽高。
# 根据用户导入的订单信息, 汇总订单明细,并计算订单内商品体积。
# 根据订单商品的体积与用户所选择品牌的箱型，推荐合适的箱子。
```

### 二、文件说明

**配置表.xlsx:**  用来记录货品的 *长宽高* 与 箱型的 *长宽高* , 必须与`箱型计算.exe`放在同一级目录下。

**箱型计算.exe:** 程序入口。



![1567141336441](https://raw.githubusercontent.com/SmallPotY/unboxing/master/README.assets/1567141336441.png)

### 三、界面说明

![1567139912936](https://raw.githubusercontent.com/SmallPotY/unboxing/master/README.assets/1567139912936.png)

![1567140354577](https://raw.githubusercontent.com/SmallPotY/unboxing/master/README.assets/1567140354577.png)



### 四、计算结果说明

![1567142201802](https://raw.githubusercontent.com/SmallPotY/unboxing/master/README.assets/1567142201802.png)

### 五、部分错误说明

##### a. 选择的数据表格式不对，表头没有对应字段

![1567140494677](https://raw.githubusercontent.com/SmallPotY/unboxing/master/README.assets/1567140494677.png)



##### b. 配置表.xlsx 不在应用程序的同级目录下, 或配置表的格式被修改，无法正确读取资料

![1567140564227](https://raw.githubusercontent.com/SmallPotY/unboxing/master/README.assets/1567140564227.png)



##### c. 运算失败，部分商品没资料的情况下未勾选 ”缺资料商品不提示“

![1567140775711](https://raw.githubusercontent.com/SmallPotY/unboxing/master/README.assets/1567140775711.png)



##### d. 预留空间必须是0~99数字, 预留空间仅能填写数字,且不得超过99

![1567140924669](https://raw.githubusercontent.com/SmallPotY/unboxing/master/README.assets/1567140924669.png)
