# poi-excel-template简介
poi-excel-template是一个基于Apache POI的Excel模板引擎，也是一个免费开源的Java类库，你可以非常方便的加入到你的项目中，根据定义的Excel模板导出你想的Excel

# 前言
最近项目上需要导出一个复杂的excel, 完全使用poi生成比较费时，也不易维护，所以本着偷懒的思维，实现了一套根据模板导出Excel的通用功能，现将它开源，与大家分享!

---

# 博客地址
[https://blog.csdn.net/scm_2008/article/details/127368510](https://blog.csdn.net/scm_2008/article/details/127368510)

---

# 原理
总体原理就是使用占位符进行文本替换。
1. 静态替换. 格式`{{key}}` 例如：在map里增加`title`的key，那么excel中所有`{{title}}`的占位符都会被文本替换成map中title对应的value。
2. 动态替换. 格式`{{rowid.key}}` , 我们只需要在excel里定义模板行这一行，生成时会根据实际rowid的`list.size()`动态生成`N`行，然后再对N行根据文本替换.

---

# 快速上手
## 1、静态替换
1. 定义一个Excel模板文件, 包括占位符`{{title}}`
   ![在这里插入图片描述](https://img-blog.csdnimg.cn/ec30ae8f812b4745b2e81e5db019bdd6.png)
2. 然后调用`ExcelTemplateUtil.buildByTemplate`即可
   ![在这里插入图片描述](https://img-blog.csdnimg.cn/b3671225531349e1bfc3f0f894a3a8d8.png)

为了达到这个效果，我们只需要构建一个`Map`： staticSource
```java
Map<String, String> staticSource = new HashMap<>();
staticSource.put("title", "poi-excel-template");
```
然后作为参数调用`ExcelTemplateUtil.buildByTemplate`
```java
InputStream resourceAsStream = SimpleDemo.class.getClassLoader().getResourceAsStream("simple-template.xlsx");
Workbook workbook = ExcelTemplateUtil.buildByTemplate(resourceAsStream, staticSource, null);
ExcelTemplateUtil.save(workbook, "D:\\simple-poi-excel-template.xlsx");
```
> 特别说明：静态替换在一个单元格内是支持放置多个占位符的，以达到通用的目的。

## 2、动态替换
1. 在静态替换的Excel模板文件基础上, 增加占位符`{{p.id}}`等，如下图
   ![在这里插入图片描述](https://img-blog.csdnimg.cn/99725c91c60346dea901cd7f431ebc63.png)
2.  程序中会动态生成相关行，如下图
    ![在这里插入图片描述](https://img-blog.csdnimg.cn/2669653aafe246d48cd092cfe2439423.png)
    为了达到这个效果，我们还需要构建一个`List`：dynamicSourceList，每个DynamicSource会有一个id和N行的`Map`，因为`DynamicSource`只有两个属性：
```java
private String id;
private List<Map<String, String>> dataList;
```
接下来我们构建一个这个`List`：
```java
int rows = 10; // 模拟10行
List<Map<String, String>> dataList = new ArrayList<>();
for (int i = 1; i <= rows; i++) {
    // 一行
    Map<String, String> rowMap = new HashMap<>();
    rowMap.put("id", "" + i);
    rowMap.put("name", "name" + i);
    rowMap.put("price", "" + (i * 100));
    rowMap.put("unit", "unit" + i);
    rowMap.put("discount", "" + i);
    rowMap.put("sellingPrice", "" + (i * 100 - 10));
    dataList.add(rowMap);
}
// 可以创建多个id，这里只创建1个示例
List<DynamicSource> dynamicSourceList = DynamicSource.createList("p", dataList);
```
然后同理，作为参数调用`ExcelTemplateUtil.buildByTemplate`

```java
InputStream resourceAsStream = DynamicDemo.class.getClassLoader().getResourceAsStream("dynamic-template.xlsx");
Workbook workbook = ExcelTemplateUtil.buildByTemplate(resourceAsStream, staticSource, dynamicSourceList);
ExcelTemplateUtil.save(workbook, "D:\\dynamic-poi-excel-template.xlsx");
```

> 特别说明：动态替换也支持一个单元格内多个占位符. 另外，还支持多个动态id.

---

# 其它说明
buildByTemplate和save分别支持不同的重载，以满足大多数场景.
![在这里插入图片描述](https://img-blog.csdnimg.cn/d86bc4ef701446f39aae3706b4610ba4.png)

---