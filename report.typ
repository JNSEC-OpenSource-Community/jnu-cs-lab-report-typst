#set page(
  paper: "a4",
  margin: (top: 1.16in, right: 1.25in, bottom: 1in, left: 1.25in),
)
#set text(
  font: "SimSun",
  size: 14pt,
  lang: "zh",
)
#show regex("[A-Za-z0-9]+") : set text(font: "Times New Roman")
#show raw: set text(font: "Consolas")
#set par(
  leading: 4pt,
  justify: false,
)

// ==================== 可编辑内容 ====================
// 表头字段
#let course_name = [数据结构实验]
#let experiment_name = [顺序表的实现和应用]
#let experiment_date = [2026.9.11]
#let class_name = [计科2601]
#let student_name = [张三]
#let student_id = [1000100000]
#let instrument_id = []
#let report_requirements = [实验报告要求   1.实验目的   2.实验要求  3.实验步骤  4.程序清单  5.运行情况  6.流程图  7.实验体会]

// 实验报告正文
#let experiment_purpose = [
  (1) 熟悉线性表的定义和基本操作；
]
#let experiment_content = [
  aaaa
]
#let program_list = [
  ```c
  #include <stdio.h>
  int main() {
    return 0;
  }
  ```
]
#let run_status = [
  #image("example_image.png", height: 30%)
]
#let reflection = [
  求求你们不要再用 docx 交报告了
]

// ==================== 版式定义 ====================
#let header-row-height = 0.64cm
#let field(body) = box(
  width: 100%,
  height: header-row-height,
)[
  #align(left + horizon)[
    #box(width: 100%, height: 10.5pt)[
      #h(0.05cm)#body
      #place(left + bottom, dy: 1pt)[
        #line(length: 100%, stroke: 0.5pt)
      ]
    ]
  ]
]
#let header-cell(body) = box(width: 100%, height: header-row-height)[#align(left + horizon)[#body]]
#let cell-box(body, height: none, width: 100%) = box(width: width, height: height)[#body]
#let centered-cell(body) = align(center + horizon)[#text(size: 10.5pt)[#body]]
#let centered-left-cell(width: 0.75cm, body) = align(center + horizon)[
  #box(width: width)[#align(left)[#text(size: 10.5pt)[#body]]]
]
#let heading(body) = text(font: "SimSun", size: 14pt, weight: "bold")[#body]

#align(center)[
  #text(font: "SimSun", size: 18pt, weight: "bold")[江南大学人工智能与计算机学院实验报告]
]
#v(0.33cm)

#let columns = (
  2.095cm,
  0.635cm, 0.635cm, 0.635cm, 0.635cm,
  0.635cm, 0.635cm, 0.635cm, 0.635cm,
  1.270cm, 0.635cm, 1.270cm, 1.905cm, 1.280cm, 2.847cm,
)
#let instrument-field-width = if instrument_id == [] { 2.0cm } else { 1.0cm }

#let body-content = [
  #pad(left: 0.20cm, right: 0.20cm, top: 0.34cm)[
    #set text(size: 10.5pt)
    #set par(leading: 4pt)

    #heading[实验目的：]

    #experiment_purpose

    #heading[实验内容：]

    #experiment_content

    #heading[程序清单：]

    #program_list

    #heading[运行情况：]

    #run_status

    #heading[实验体会：]

    #reflection
  ]
]
#let body-min-height = 16.174cm

#table(
  columns: columns,
  stroke: 0.5pt,
  inset: 0pt,

  // 表头信息区
  table.cell(colspan: 15)[
    #box(width: 100%, height: 2.1cm)[
      #pad(left: 0.20cm, right: 0.20cm, top: 0.20cm)[
        #set text(size: 10.5pt)
        #set par(leading: 10pt)
        #table(
          columns: (1.50cm, 2.35cm, 0.12cm, 1.50cm, 3.55cm, 0.12cm, 1.50cm, 2.0cm, 0.12cm, 1.50cm, instrument-field-width),
          stroke: none,
          inset: 0pt,
          align: left,
          [#header-cell[课程名称]], [#header-cell[#field[#course_name]]],
          [],
          [#header-cell[实验名称]], [#header-cell[#field[#experiment_name]]],
          [],
          [#header-cell[实验日期]], [#header-cell[#field[#experiment_date]]],
          table.cell(colspan: 3)[#header-cell[]],
          [#header-cell[班#h(0.75cm)级]], [#header-cell[#field[#class_name]]],
          [],
          [#header-cell[姓#h(0.75cm)名]], [#header-cell[#field[#student_name]]],
          [],
          [#header-cell[学#h(0.75cm)号]], [#header-cell[#field[#student_id]]],
          [],
          [#header-cell[仪器编号]], [#header-cell[#field[#instrument_id]]],
        )
        #v(-0.08cm)
        #text(size: 9pt)[#report_requirements]
      ]
    ]
    ],

  // 主体填写区
  table.cell(colspan: 15)[
    #context {
      let body-size = measure(body-content)
      let extra-height = calc.max(0pt, body-min-height - body-size.height)
      block[#body-content #v(extra-height)]
    }
  ],

  // 教师评价区
  [#cell-box(height: 1.371cm)[#centered-cell[教师评价]]],
  [#cell-box(height: 1.371cm)[#centered-cell[优]]],
  [#cell-box(height: 1.371cm)[]],
  [#cell-box(height: 1.371cm)[#centered-cell[良]]],
  [#cell-box(height: 1.371cm)[]],
  [#cell-box(height: 1.371cm)[#centered-cell[中]]],
  [#cell-box(height: 1.371cm)[]],
  [#cell-box(height: 1.371cm)[#centered-cell[及#linebreak()格]]],
  [#cell-box(height: 1.371cm)[]],
  [#cell-box(height: 1.371cm)[#centered-left-cell(width: 0.75cm)[不及#linebreak()格]]],
  [#cell-box(height: 1.371cm)[]],
  [#cell-box(height: 1.371cm)[#centered-cell[教师#linebreak()签名]]],
  [#cell-box(height: 1.371cm)[]],
  [#cell-box(height: 1.371cm)[#centered-cell[日期]]],
  [#cell-box(height: 1.371cm)[]],
)
