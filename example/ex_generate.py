from genquiz.generator import *

# load students list
students = Student('./student.xlsx')

for student_id, student_name in students:
    # question 1
    text1 = Text('1.(10 คะแนน) จงหาตัวอย่างที่{method}กับตัวอย่างที่ 1 ที่สุดด้วยวิธีการหาระยะห่างด้วยวิธี{distance}',
                 {'method': ['เหมือน', 'ต่าง'],
                  'distance': ['ยูคลิด', 'แมนฮัตตัน']})
    table1 = Table({'ตัวอย่างที่': [1, 2, 3, 4, 5],
                    'อายุ (ปี)': np.random.randint(10, 50, 5),
                    'ความสูง (ซม.)': np.random.randint(100, 200, 5),
                    'น้ำหนัก (กก.)': np.random.randint(50, 100, 5)})
    # find solution of question 1
    d = table1.df.iloc[1:, 1:].values - table1.df.iloc[0, 1:].values
    if text1.random_list['distance'] == 'ยูคลิด':
        d = np.sqrt(np.sum(d ** 2, axis=1))
    if text1.random_list['distance'] == 'แมนฮัตตัน':
        d = np.sum(np.abs(d), axis=0)
    if text1.random_list['method'] == 'เหมือน':
        i = d.argmin()
    if text1.random_list['method'] == 'ต่าง':
        i = d.argmax()
    sol = np.array([i + 2, d[i]])
    sol = sol.reshape(2, 1)
    solution1 = Solution(sol, 'B2:B3', score=10)
    question1 = [text1, table1]

    # question 2
    text2 = Text('2.(10 คะแนน) จงแก้สมการต่อไปนี้ (โดยให้ตอบคำถามเรียงจากน้อยไปมาก)')
    question2 = [text2]
    x, y = sympy.symbols('x y')
    sol = []
    for i in range(10):
        a, b = np.random.randint(-10, 10, 2)
        question2.append(Equation(x ** 2 + (a + b) * x + (a * b), 0))
        sol.append(sorted([-a, -b]))  # sort for grading
    sol = np.array(sol)
    solution2 = Solution(sol, 'C6:D15', score=10)

    # question 3
    text3 = Text('3.(5 คะแนน) จงหาพื้นที่แรงเงา (ทศนิยม 2 ตำแหน่ง)')
    figure3 = Figure('./figures', width=4)  # 4 inches width
    question3 = [text3, figure3]
    solution3 = Solution(figure3.sol, 'B18', score=5)

    # make quiz
    Quiz(student_id, student_name,
         [question1,
          question2,
          question3],
         [solution1,
          solution2,
          solution3],
         question_path='./exam1',
         solution_path='./sol1'
         )
