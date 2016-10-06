using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using System.Collections;

namespace DKKReport
{
    class SimpleItem
    {
        string _name;
        double _valuePeriod;
        double _valueYear;

        public string Name
        {
            get { return _name; }
            set { _name = value; }
        }
        public double ValuePeriod
        {
            get { return _valuePeriod; }
            set { _valuePeriod = Math.Abs ( value ); }
        }
        public double ValueYear
        {
            get { return _valueYear; }
            set { _valueYear = Math.Abs ( value ); }
        }
        public SimpleItem ( string name, string valuep, string valuey )
        {
            Name = name;
            ValuePeriod = double.Parse ( valuep );
            ValueYear = double.Parse ( valuey );
        }
        public void PrintData ( Excel.Worksheet ws, ref long rLastRow, long k )
        {
            ws.Cells[rLastRow, 1].value = k;
            ws.Cells[rLastRow, 2].value = Name;
            ws.Cells[rLastRow, 3].value = ValuePeriod;
            ws.Cells[rLastRow, 4].value = ValueYear;
            //ws.Cells[rLastRow, 3].NumberFormat = "#,##0.00";
            //ws.Cells[rLastRow, 4].NumberFormat = "#,##0.00";
            rLastRow++;
            
        }
    }

    class ComplexItem
    {
        string _name;
        long _valuePeriod;
        long _valueYear;
        string _parentName;
        long _level;
        public static long rLastRow; //статическая переменная  - последняя строка
        public static int rSeconLevelCount; // стат переменная - обнуляеся при переход к новому глобальному компелксу (доход - расход). а внутри - при прибавлении комплексов второго уровня увеличивается на 1 чтобы можно было пронумеровать каждый из них. Грубо говоря - счетчик сколько комплексов второго уровня в данном глобальном.

        long _colorIndex;

        public string del;
        public List<SimpleItem> LSimpleItem = new List<SimpleItem> ();
        public List<ComplexItem> LComplexItem = new List<ComplexItem> ();

        public Hashtable LHashComplexItem = new Hashtable();
        

        public string Name
        {
            get { return _name; }
            set { _name = value; }
        }
        public long ValuePeriod
        {
            get { return _valuePeriod; }
            set { _valuePeriod = value; }
        }
        public long ValueYear
        {
            get { return _valueYear; }
            set { _valueYear = value; }
        }
        public string ParentName
        {
            get { return _parentName; }
            set { _parentName = value; }
        }
        public long Level
        {
            get { return _level; }
            set { _level = value; }
        }
        public void AddSimleItem ( SimpleItem sItem )
        {
            LSimpleItem.Add ( sItem );
        }
        public void AddComplexItem ( ComplexItem cItem )
        {
            LComplexItem.Add ( cItem );
            LHashComplexItem.Add ( cItem.Name, cItem );
        }

        public ComplexItem (  )
        {
            this.Name = "temp";
        }
        public ComplexItem ( string n, string pName )
        {
            this.Name = n;
            this.ParentName = pName;
        }
        public ComplexItem ( string n, string pName, string plev )
        {
            this.Name = n;
            this.ParentName = pName;
            this.Level = long.Parse ( plev );
            switch (int.Parse ( plev )) //при этом конструкторе присваиваестся значение цвета фона (если первый уровень - синий , если второе -желтый )
            {
                case 1:
                    _colorIndex = 34;
                    break;
                case 2:
                    _colorIndex = 6;
                    break;
            }
        }
        public ComplexItem GetCItem ( string tname )
        {
           
            //foreach (ComplexItem c in LComplexItem)
            //{
            //    if (c.Name == tname)
            //    {
            //        itogItem = c;
            //        break;
            //    }

            //}
            ComplexItem itogItem = (ComplexItem)LHashComplexItem[tname];
            //if (itogItem==null)
            //      itogItem = new ComplexItem ( "temp", "temp" );
            return itogItem;
            //ComplexItem itogItem =(ComplexItem) LHashComplexItem[tname];
            //return itogItem;
        }

        

        public long PrintData ( Excel.Worksheet ws ) //вывод на лист
        {
            long rItog=777;
            ws.Cells[rLastRow, 2].value = Name; //имя статьи во 2 столбец
            ws.Range[ws.Cells[rLastRow, 2], ws.Cells[rLastRow, 4]].interior.colorindex = _colorIndex; //покарска цветом
            ws.Range[ws.Cells[rLastRow, 1], ws.Cells[rLastRow, 4]].font.bold = true; 
            if (Level == 2) //если уровень второй (первый это доходы и расходы) - то номер выставляем римскими числами
            {
                List<string> arab = new List<string> ();
                arab.Add ( "I" );
                arab.Add ( "II" );
                arab.Add ( "III" );
                arab.Add ( "IV" );
                arab.Add ( "V" );
                arab.Add ( "VI" );
               // arab.Add ( "VI" );
                arab.Add ( "VII" );
                arab.Add ( "VIII" );
                arab.Add ( "IX" );
                arab.Add ( "X" );
                arab.Add ( "XII" );
                arab.Add ( "XIII" );
                arab.Add ( "XIV" );
                arab.Add ( "XV" );
                arab.Add ( "XVI" );
                arab.Add ( "XVII" );
                arab.Add ( "XVIII" );
                arab.Add ( "XIX" );
                arab.Add ( "XX" );

                ws.Cells[rLastRow, 1].value = arab[ComplexItem.rSeconLevelCount++]; // и увеличиваем переменную на 1
            }
            if (LSimpleItem.Count > 0)// если в статье сразу есть пункты расходов - то выводим их (!Важно - даже если там есть еще комплексы - их вывод игнорится)
            {
                if (LSimpleItem.Count == 1 && LSimpleItem[0].Name==this.Name) //если всего один пункт расходов и его имя совпадает с названием статьи - его не выводим
                {
                    
                    ws.Cells[rLastRow, 3].value = LSimpleItem[0].ValuePeriod;
                    ws.Cells[rLastRow, 4].value = LSimpleItem[0].ValueYear;
                    rItog = rLastRow++;
                    rLastRow++;
                }
                else//если пунктов больше
                {
                    rLastRow++;
                    long r1;
                    long r2;
                    long k = 1;
                    r1 = rLastRow;
                    foreach (SimpleItem si in LSimpleItem) //каждый симпл выводим
                    {
                        si.PrintData ( ws, ref rLastRow, k++ );
                    }
                    r2 = rLastRow;
                    r2--;
                    ws.Cells[rLastRow, 2].value = "Итого по группе: " + Name; //в конце пишем итог
                    //  ws.Range[ws.Cells[rLastRow, 2], ws.Cells[rLastRow, 4]].interior.colorindex = _colorIndex;
                    ws.Range[ws.Cells[rLastRow, 2], ws.Cells[rLastRow, 4]].font.bold = true;
                    //ws.Cells[rLastRow++, 3].value = r1 + "-" + r2;
                    rItog = rLastRow;
                    ws.Cells[rLastRow, 3].formulaR1C1 = "=SUM(R" + r1 + "C:R" + r2 + "C)"; //и формулы итогов для периода и года
                    ws.Cells[rLastRow++, 4].formulaR1C1 = "=SUM(R" + r1 + "C:R" + r2 + "C)";
                    rLastRow++;
                }
            }
            else//  если нету симплов)
            {

                if (LComplexItem.Count > 0) // и есть комплексы
                {
                    rLastRow++;
                    List<string> rList = new List<string> ();

                    foreach (ComplexItem c in LComplexItem) //для каждого комплекса
                    {
                        //if (Name == "Имущество,Капит. влож. Инвестиции")
                        //    testdel = 5;

                        long rItogComplex = c .PrintData ( ws ); //комплекс выводится на лист (рекурсия, ага) 
                        rList.Add ( "R" + rItogComplex + "C" );//и возвращается итоговая строка этого комлекса
                        //if (Name == "Имущество,Капит. влож. Инвестиции")
                        //    testdel = 5;
                    }

                    //   string f2 = "=";
                    //for (int i = 0; i < rList.Count; i++)
                    //{
                    //    f1 = f1 +"+" + rList[i];

                    //}

                    string f1 = "0";
                    if (rList.Count > 0) //когда все комплексы принадлежащие этому комплексу пройдены - подбивается строка итога вида =r15c3+r25c3...
                        f1 = "=" + string.Join ( "+", rList );
                    rItog = rLastRow;
                    ws.Cells[rLastRow, 2].value = "Итого по группе: " + Name;
                  //  ws.Range[ws.Cells[rLastRow, 2], ws.Cells[rLastRow, 4]].interior.colorindex = _colorIndex;
                    ws.Range[ws.Cells[rLastRow, 2], ws.Cells[rLastRow, 4]].font.bold = true;
                    ws.Cells[rLastRow, 3].formular1c1 = f1;
                    ws.Cells[rLastRow++, 4].formular1c1 = f1;
                    rLastRow++;
                }
                else //и если нет не комплексов не симплов
                {
                    rItog = rLastRow;
                    ws.Cells[rLastRow, 3].value = 0;
                    ws.Cells[rLastRow++, 4].value = 0;
                    
                    rLastRow++;
                }
            }
            //if (rItog == 777)
            //{
            //    long k = 1;
            //}
            return rItog;
        }
        public override string ToString() {
            return _name;
        }
    }

    
}
