using System;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;



namespace DepartmentEmployee
{
    public partial class Form1 : Form
    {

        MSSQLConnectionParams mssql_params = new MSSQLConnectionParams();
        DB db;

        public Form1()
        {
            InitializeComponent();

            //!!!!!!!!!!УЧЕТНЫЕ ДАННЫЕ MS SQL!!!!!!!!!!!!!!!
            mssql_params.hostname = "SERVER2012R2";     //Hostname
            mssql_params.database = "TestDB";          //Database
            mssql_params.user = "sa";                   //Username
            mssql_params.passwd = "dgfdsfgsdfgsdfds";       //Pass
            //!!!!!!!!!!УЧЕТНЫЕ ДАННЫЕ MS SQL!!!!!!!!!!!!!!!


            db = new DB(mssql_params);



            RefreshDeparts(); // Обновляем список департаментов.
        }



        // ======================Обновлене Departments===========
        //=================Обновление корня Departments=================
        private async void RefreshDeparts()
        {
         

            //Очищаем treeview на случай если там чтото есть
            TreeView1.Nodes.Clear();


            //Этим этапом (всей этой функцией рефреша) мы подгружаем только корневые ноды. Тоесть те у которых ParentID = null, остальные (внутренние) мы подгружаем с помощью обработки события выделения ноды TreeView1_AfterSelect.

            //Получаем datatable из соответвующей функции - получаем только корневые папки - где парект ID = 0
            DataTable dt_tt = await db.GetDatatableFromMSSQLAsync("SELECT ID, ParentDepartmentID, Code, Name FROM Department WHERE ParentDepartmentID IS NULL");
            

            // для каждого элемента в datatable
            foreach (DataRow row in dt_tt.Rows)
            {
                //обявляем переменные с ячейками текущей перебираемой строки
                string current_row_id = row["ID"].ToString();
                string current_row__parent_id = row["ParentDepartmentID"].ToString();
                string current_row__name = row["Name"].ToString();
                string current_row__code = row["Code"].ToString();
                
                TreeNode node = new TreeNode(current_row__name + " (" + current_row__code + ")");
                //Присваем Tag этому элементу равный текущему ID                    
                node.Tag = current_row_id;
                //Добавляем его в treeview
                TreeView1.Nodes.Add(node);


            }

            // Пройдемся по всем корневым нодам выделением, чтобы сработало событие выделения ноды и подгрузились дочерние элементы
            foreach (TreeNode node in TreeView1.Nodes)
            {
                TreeView1.SelectedNode = node;
            }

            //Раскроем все свернутые ноды
            TreeView1.ExpandAll();



            //=================Конец Обновление корня departments=================
        }


        // Событие при выборе элемента, в нем мы снова перебираем datatable в поисках элементов parentID которых соответвует TagID пункта меню от которого сработало это событие
        private void TreeView1_AfterSelect(object sender, TreeViewEventArgs e)
        {

            //Получаем datatable
            DataTable dt_tt = db.GetDatatableFromMSSQL("SELECT ID, Code, Name FROM Department WHERE ParentDepartmentID = '" + e.Node.Tag.ToString() + "'");


            // Перебираем строки в таблице datatable
            foreach (DataRow row in dt_tt.Rows) {
                // обявляем переменные с ячейками текущей перебираемой строки

                string current_row_id = row["ID"].ToString();
                string current_row__name = row["Name"].ToString();
                string current_row__code = row["Code"].ToString();

                // Объявляем триггер который с помощью которого будем проверять не добавлены ли УЖЕ добавляемые элементы в меню при прошлом клике на него
                bool already_added = false;

                // Если parentid в текущей строке datatable = Tag текущего элемента от которого сработало событие                

                // Проверяем что уже не добавили эти элемнеты меню при прошлых кликах на пункте меню
                // Перебираем дочерние элементы текушего пункта ментю от котого сработало событие
                foreach (TreeNode mynode in TreeView1.SelectedNode.Nodes)
                {
                    // Если хотя бы у одного элемента Tag = CurrentRowID значит мы уже обрабатывали этот пункт меню ранее, следовательно включаем триггер
                    if (mynode.Tag.ToString() == current_row_id)
                    {
                        already_added = true;
                    }
                }
                // Если триггер того что меню было обработано ранее всё еще равен = false
                if (already_added == false) {
                    //Создаем объект Treenode с именем из текущей строчки перебираемой datatable
                    TreeNode node = new TreeNode(current_row__name + " (" + current_row__code + ")");
                    //Присваиваем ему Tag  из текущей строчки перебираемой datatable
                    node.Tag = current_row_id;
                    //Добавляем его к выбраному пункту меню от которого сработало это событие.
                    TreeView1.SelectedNode.Nodes.Add(node);

                }
            }
            // Развернем выделеную ноду
            TreeView1.SelectedNode.Expand();

            refreshGrid();

        }



        private async void refreshGrid()
        {            
            try
            {
                    string current_department_id = TreeView1.SelectedNode.Tag.ToString();
                    DataTable dt = await db.GetDatatableFromMSSQLAsync("SELECT ID, DepartmentID, SurName, FirstName, Patronymic, DateOfBirth, DocSeries, DocNumber, Position FROM Empoyee WHERE DepartmentID = '" + current_department_id + "'");



                    //Добавим колонку Age к DataTable и посчитаем возраст сотрудника
                    dt.Columns.Add("Age");                   
                      for (int i = 0; i < dt.Rows.Count; i++)
                      {
                            //Посчитаем разницу между текущей датой и датой рождения из соответвующей ячейки
                            TimeSpan DateTimeDifference = DateTime.Now - (DateTime)dt.Rows[i]["DateOfBirth"];

                            //Разницу в днях делим на 365 и округляем
                            dt.Rows[i]["Age"] = Math.Round(DateTimeDifference.TotalDays / 365, 0);
                      }


                DataGridView1.DataSource = dt; //Присвеиваем DataTable в качестве источника данных DataGridView
     


            }
            catch
            {                

                DataTable dt = await db.GetDatatableFromMSSQLAsync("SELECT ID, DepartmentID, SurName, FirstName, Patronymic, DateOfBirth, DocSeries, DocNumber, Position FROM Empoyee");
                DataGridView1.DataSource = dt;
            }



            try
            {
            // Скроем столбец ненужные столбцы
            DataGridView1.Columns["ID"].Visible = false;
            DataGridView1.Columns["DepartmentID"].Visible = false;
           

            //Заголовки таблицы
            DataGridView1.Columns["SurName"].HeaderText = "Фамилия";
            DataGridView1.Columns["FirstName"].HeaderText = "Имя";
            DataGridView1.Columns["Patronymic"].HeaderText = "Отчество";
            DataGridView1.Columns["DateOfBirth"].HeaderText = "Дата рождения";
            DataGridView1.Columns["DocSeries"].HeaderText = "Серия документа";
            DataGridView1.Columns["DocNumber"].HeaderText = "Номер документа";
            DataGridView1.Columns["Position"].HeaderText = "Должность";
            DataGridView1.Columns["Age"].HeaderText = "Возраст";
            }
            catch { }

            // выбираем первую строчку
            try
            {
                DataGridView1.CurrentCell = DataGridView1.Rows[0].Cells[0];
            }
            catch { }

        }

        // Метод редактирования текщей ноды treeview - аргументом принимаем ноду, которую и требуется редактировать
        private async void EditCurrentDepartment(TreeNode node)
        {
            Form2 departments_form = new Form2();

            

            departments_form.TextBox1.Text = node.Text.Substring(0, node.Text.LastIndexOf("(") - 1);
            departments_form.textBox2.Text = (node.Text.Split('(').Last()).Substring(0, (node.Text.Split('(').Last()).Length-1);

            if (departments_form.ShowDialog() != DialogResult.OK)
            { return; }

            Guid id = Guid.NewGuid();

            string newname = departments_form.TextBox1.Text.Replace("'", "''");
            string newcode = departments_form.textBox2.Text.Replace("'", "''");

            // Объявляем переменную для имени папки


            //Если пользователь ничего не ввел выходим из процедуры
            if (newname == "") { return; }


            string tag = node.Tag.ToString();
            
            bool sqlresult = await db.ExecSQLAsync("UPDATE Department set Name='" + newname + "',Code='" + newcode + "'  WHERE ID = '" + tag + "'");
            node.Text = (newname + " (" + newcode + ")");

        }

        //Кнопка онбновления Departments
        private void Button2_Click(object sender, EventArgs e)
        {
            RefreshDeparts();
        }

        //Кнопка редактирование Departments
        private void Button30_Click(object sender, EventArgs e)
        {
            //Если не выбран элемент который мы собираемся удалять выходим
            if (TreeView1.SelectedNode == null)
            {
                return;
            }

            //Запускаем функцию редактирования папки и передаем её текущую ноду которую будем редактировать
            EditCurrentDepartment(TreeView1.SelectedNode);
        }

        //Кнопка добавления нового Department
        private async void Button28_Click(object sender, EventArgs e)
        {
            Form2 departments_form = new Form2();
            if (departments_form.ShowDialog() != DialogResult.OK)
            { return;}


            Guid id = Guid.NewGuid(); 
            string name = departments_form.TextBox1.Text.Replace("'", "''");
            string tag = "";
            string code = departments_form.textBox2.Text.Replace("'", "''");


            //Если пользователь ничего не ввел выходим
            if (name == "")
            { return; }


            //Записываем в переменную Tag, Tag текущего выбраного элемента (это его ID в базе, которое мы будем использовать чтобы задать родителя нашего нового добавляемого элемента)
          
            if(TreeView1.SelectedNode == null)
              {tag = "";}

             else
              {tag = TreeView1.SelectedNode.Tag.ToString();}


            // Записываем в базу новую строчку, задав поля имя и parent ID

            bool sqlresult = await db.ExecSQLAsync("INSERT into Department(ID, ParentDepartmentID, Name, Code) VALUES('" + id + "','" + tag + "', '" + name + "', '" + code + "')");


            // Получаем данные обратно из базы

            //Получаем datatable из функции 
            //сразу выборка из таблицы - получаем последнюю строчку - она и есть та которую мы только что добавили
            DataTable dt_tt = await db.GetDatatableFromMSSQLAsync("SELECT ID, ParentDepartmentID, Name, Code FROM Department WHERE ID = '" + id + "'");


            //Находим последнюю запись в таблице, - она и есть та которую мы только что добавили
            //Получаем индекс последней записи в нашем массиве
            int LastRowIndex = dt_tt.Rows.Count - 1;

            // Получаем последнюю строчку по этому индексу
            DataRow lastrow = dt_tt.Rows[LastRowIndex];

            // Задаем переменные для столбцов этой строчки
            string lastrow_ID = lastrow["ID"].ToString();
            string lastrow_ParentID = lastrow["ParentDepartmentID"].ToString();
            string lastrow_Name = lastrow["Name"].ToString();
            string lastrow_Code = lastrow["Code"].ToString();



            //Тут будем добавлять это в наш treeview

            TreeNode node = new TreeNode((lastrow_Name + " (" + lastrow_Code + ")"));  //созадем объект node с именем из базы
            node.Tag = lastrow_ID; //присваиваем ему tag, который соответвует его ID в базе

        //Если ниче не выделено то добавляем ноду в корень treeview
            if (TreeView1.SelectedNode == null)
            {
                TreeView1.Nodes.Add(node);
            // А если выделено то добавляем в то место которое выделено 
             }
            else
            {
                TreeView1.SelectedNode.Nodes.Add(node);
                TreeView1.SelectedNode.Expand(); // И разворачиваем 
            }
        }



        



        //Перехват отрисовки нод в Treeview1
        private void TreeView1_DrawNode(object sender, DrawTreeNodeEventArgs e)
        {


            SolidBrush NodeBursh = new SolidBrush(System.Drawing.Color.White);
            SolidBrush SlectedNodeBursh = new SolidBrush(System.Drawing.Color.FromArgb(185, 209, 234));
            System.Drawing.Color textForeColor = e.Node.ForeColor;



            if (e.Node.IsSelected)
            {
                if (TreeView1.Focused)
                {
                    e.Graphics.FillRectangle(SlectedNodeBursh, e.Bounds);
                }
                TextRenderer.DrawText(e.Graphics, e.Node.Text, e.Node.TreeView.Font, e.Node.Bounds, textForeColor);
            }
            else
            {
                e.Graphics.FillRectangle(NodeBursh, e.Bounds);

                TextRenderer.DrawText(e.Graphics, e.Node.Text, e.Node.TreeView.Font, e.Node.Bounds, textForeColor);
            }

        }

        //Кнопка удаления Departments
        private async void Button29_Click(object sender, EventArgs e)
        {
            //Если не выбран элемент который мы собираемся удалять выходим
            if (TreeView1.SelectedNode == null)
            {
                MessageBox.Show("Сначала выберите элемент.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1);
                return;
            }

            string current_department_id = TreeView1.SelectedNode.Tag.ToString();


            



            //Если внутри выбраного элемента содержаться другие элементы выводим уведомление и выходим из процедуры
            if (TreeView1.SelectedNode.Nodes.Count != 0)
            {
                MessageBox.Show("Нельзя удалить элемент, поскольку он содержит вложенные элементы", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1);
                return;
            }


            //Проверим существуют ли сотрудники, привязанные к данной таблице
            DataTable dt_tt_employee = await db.GetDatatableFromMSSQLAsync("select ID from Empoyee where DepartmentID = '" + current_department_id + "'");
            if (dt_tt_employee.Rows.Count > 0)
            {
                MessageBox.Show("Нельзя удалить элемент, поскольку есть сотрудники привязанные к данному элементу", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1);
                return;
            }


            //Получаем DataTable
            //сразу в запросе находим нужные строчки и добавляем в datatable посути одну нужнную строку

            DataTable dt_tt = await db.GetDatatableFromMSSQLAsync("SELECT ID FROM Department WHERE ID = '" + current_department_id + "'");            
            // Перебираем все строчки в таблице

            foreach (DataRow row in dt_tt.Rows)
            {
                // объявляем переменную которая с ID шникном текущей строчки в таблице
                string row_ID = row["ID"].ToString();
                
                //Если этот ID равен текущему Tag выбраного элемента, то удаляем эту строчку из базы

                if (row_ID == TreeView1.SelectedNode.Tag.ToString())
                {
                    bool sqlresult = await db.ExecSQLAsync("DELETE FROM Department WHERE ID = '" + row_ID + "'");
                }

            }

            // А потом удаляем эту ноду из Treeview
            TreeView1.SelectedNode.Remove();

        }


        //Кнопка добавить нового сотрудника
        private async void Button4_Click(object sender, EventArgs e)
        {

            Form3 employee_form = new Form3();

            if (employee_form.ShowDialog() != DialogResult.OK)
            { return; }

            string FirstName = employee_form.textBox1.Text.Replace("'", "''");
            string SurName = employee_form.textBox2.Text.Replace("'", "''");
            string Patronymic = employee_form.textBox3.Text.Replace("'", "''");
            DateTime DateOfBirth = employee_form.dateTimePicker1.Value.Date;
            string DocSeries = employee_form.textBox4.Text.Replace("'", "''");
            string DocNumber = employee_form.textBox5.Text.Replace("'", "''");
            string Position = employee_form.textBox6.Text.Replace("'", "''");
            
            string DepartmentID = TreeView1.SelectedNode.Tag.ToString();
            
            


            //записываем данные из текстбоксов Form3 в наши переменные
            // А потом экранируем кавычечку

            bool sqlresult = await db.ExecSQLAsync("INSERT into Empoyee(FirstName, SurName, Patronymic, DateOfBirth, DocSeries, DocNumber, Position, DepartmentID) values('" + FirstName + "', '" + SurName + "', '" + Patronymic + "', '" + DateOfBirth + "', '" + DocSeries + "', '" + DocNumber + "', '" + Position + "', '" + DepartmentID + "')");

            refreshGrid();

        }

        //Кнопка редактировать сотрудников
        private async void Button7_Click(object sender, EventArgs e)
        {
            string ID = "";
            string FirstName;
            string SurName;
            string Patronymic;
            DateTime DateOfBirth;
            string DocSeries;
            string DocNumber;
            string Position;


            try
            {
                ID = DataGridView1.CurrentRow.Cells["ID"].Value.ToString();
            }
            catch {
                MessageBox.Show("Сначала выберите сотрудника", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1);
              return;
            }


            Form3 employee_form = new Form3();

            //Заполняем в Form2 поля для того чтобы было что редактировать.
            employee_form.textBox1.Text = DataGridView1.CurrentRow.Cells["FirstName"].Value.ToString();
            employee_form.textBox2.Text = DataGridView1.CurrentRow.Cells["SurName"].Value.ToString();
            employee_form.textBox3.Text = DataGridView1.CurrentRow.Cells["Patronymic"].Value.ToString();
            employee_form.dateTimePicker1.Value = (DateTime)DataGridView1.CurrentRow.Cells["DateOfBirth"].Value;
            employee_form.textBox4.Text = DataGridView1.CurrentRow.Cells["DocSeries"].Value.ToString();
            employee_form.textBox5.Text = DataGridView1.CurrentRow.Cells["DocNumber"].Value.ToString();
            employee_form.textBox6.Text = DataGridView1.CurrentRow.Cells["Position"].Value.ToString();
           


            if (employee_form.ShowDialog() != DialogResult.OK)
            { return; }



           
            FirstName = employee_form.textBox1.Text.Replace("'", "''");
            SurName = employee_form.textBox2.Text.Replace("'", "''");
            Patronymic = employee_form.textBox3.Text.Replace("'", "''");
            DateOfBirth = employee_form.dateTimePicker1.Value.Date;
            DocSeries = employee_form.textBox4.Text.Replace("'", "''");
            DocNumber = employee_form.textBox5.Text.Replace("'", "''");
            Position = employee_form.textBox6.Text.Replace("'", "''");


            bool sqlresult = await db.ExecSQLAsync("UPDATE Empoyee set FirstName='" + FirstName + "',SurName= '" + SurName + "',Patronymic= '" + Patronymic + "',DateOfBirth= '" + DateOfBirth + "',DocSeries= '" + DocSeries + "',DocNumber= '" + DocNumber + "',Position= '" + Position + "' where ID = '" + ID + "'");



            refreshGrid();
        }

        //Кнопка удаления сотрудника
        private async void Button3_Click(object sender, EventArgs e)
        {
            string ID = "";

            try
            {
                ID = DataGridView1.CurrentRow.Cells["ID"].Value.ToString();
            }
            catch
            {
                MessageBox.Show("Сначала выберите сотрудника", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1);
                return;
            }

            

            //Удаляем из базы
        if ((DialogResult = MessageBox.Show("Вы действительно хотите удалить сотрудника: " + DataGridView1.CurrentRow.Cells["FirstName"].Value + " " + DataGridView1.CurrentRow.Cells["SurName"].Value + "?", "Delete Employee", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1)) == DialogResult.Yes)
            {
                bool sqlresult = await db.ExecSQLAsync("DELETE FROM Empoyee where ID = '" + ID + "'");
            }

            //Удаляем из DataGridView
            DataGridView1.Rows.Remove(DataGridView1.CurrentRow); 

        }

        //Обработчик события даблклика по ячейке DatagridView
        private void DataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            Button7_Click(sender, e);
        }

        //Обработчик события даблклика по Threeview Node
        private void TreeView1_NodeMouseDoubleClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            EditCurrentDepartment(e.Node);
        }





        // ====Обработка DragAndDrop для перемещения сотрудников между отделами путем перетаскивания

        // Событие срабатывает при движении мышкой по области datagridview. Если нажата левая кнопка мышки то захватываем элемент для drag and drop а
        private void DataGridView1_MouseMove(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                try
                {
                    string dragedItemID = DataGridView1.CurrentRow.Cells["ID"].Value.ToString();
                    DataGridView1.DoDragDrop(dragedItemID, DragDropEffects.Move);
                }
                catch
                {
                }
            }                    
        }

        //Событие вхождения перетаскиваемого объекта в зону ThreeView
        private void TreeView1_DragEnter(object sender, DragEventArgs e)
        {
            e.Effect = DragDropEffects.Move; //Курсор DragAndDrop
        }


        //Обработка события перетаскиваниея
        private async void TreeView1_DragDrop(object sender, DragEventArgs e)
        {
            //Перетаскиваемый элемент - это ID типа string - проверим что это действительно он;
            if (e.Data.GetDataPresent("System.String", false))
            {


                TreeView treeview = (TreeView)sender;                 
                TreeNode DestinationNode = null;
                Point pt;
                // находим координаты того места куда сбросили перетаскиваемый объект
                pt = treeview.PointToClient(new Point(e.X, e.Y));
                // Находим ноду, которая находилась в этих координатах
                DestinationNode = treeview.GetNodeAt(pt);
                // Проверяем что там вообще есть нода в этом месте
                if (DestinationNode == null)
                {
                    return;
                }


                // Обновляем запись в таблице Empoyee - присваиваем полю DepartmentID - Tag нашей ноды

                bool sqlresult = await db.ExecSQLAsync("UPDATE Empoyee set DepartmentID='" + DestinationNode.Tag + "' where ID = '" + e.Data.GetData("System.String") + "'");
                


                TreeView1.SelectedNode = DestinationNode; //Выбираем ноду, в которую перетащили объект


            }


        }

        //Кнопка добавления новой фирмы
        private async void button1_Click(object sender, EventArgs e)
        {
            Form2 departments_form = new Form2();
            if (departments_form.ShowDialog() != DialogResult.OK)
            { return; }


            Guid id = Guid.NewGuid();
            string name = departments_form.TextBox1.Text.Replace("'", "''");
            string code = departments_form.textBox2.Text.Replace("'", "''");


            //Если пользователь ничего не ввел выходим
            if (name == "")
            { return; }


            

            // Записываем в базу новую строчку, задав поля имя и parent ID

            bool sqlresult = await db.ExecSQLAsync("INSERT into Department(ID, Name, Code) VALUES('" + id + "', '" + name + "', '" + code + "')");


            // Получаем данные обратно из базы

            //Получаем datatable из функции 
            //сразу выборка из таблицы - получаем последнюю строчку - она и есть та которую мы только что добавили
            DataTable dt_tt = await db.GetDatatableFromMSSQLAsync("SELECT ID, ParentDepartmentID, Name, Code FROM Department WHERE ID = '" + id + "'");


            //Находим последнюю запись в таблице, - она и есть та которую мы только что добавили
            //Получаем индекс последней записи в нашем массиве
            int LastRowIndex = dt_tt.Rows.Count - 1;

            // Получаем последнюю строчку по этому индексу
            DataRow lastrow = dt_tt.Rows[LastRowIndex];

            // Задаем переменные для столбцов этой строчки
            string lastrow_ID = lastrow["ID"].ToString();
            string lastrow_ParentID = lastrow["ParentDepartmentID"].ToString();
            string lastrow_Name = lastrow["Name"].ToString();
            string lastrow_Code = lastrow["Code"].ToString();



            //Тут будем добавлять это в наш treeview

            TreeNode node = new TreeNode((lastrow_Name + " (" + lastrow_Code + ")"));  //созадем объект node с именем из базы
            node.Tag = lastrow_ID; //присваиваем ему tag, который соответвует его ID в базе

            //добавляем ноду в корень treeview
          
             TreeView1.Nodes.Add(node);
           
      
        }

        //Обработчик нажатия клавиши Enter на Datagridview1 - Это отключение перехода на сл строку при нажатии Enter.
        private void DataGridView1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.Handled = true;
            }

        }

    
        // Обработчик нажатия клавиши Enter на Datagridview1
        private void DataGridView1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == Convert.ToChar(Keys.Enter))
            {
                Button7_Click(sender, e);
            }
        
        }

      
    }
}
