# Lecture 6. Database applications. Forms

<h3>Abstract</h3>

<p>Every database application consists of at least two parts:

<ol>
<li><i>the database</i> (<i>the server side</i>)
	which stores large amounts of data and provides fast access
	to that data (This is <i>the server side</i>) and
<li><i>the graphical user interface</i> = <i>GUI</i> (<i>the client side</i>)
	which provides functions for the users. 
	These functions include operations concerned with storing data and
	performing actions on the stored data.
</ol> 

<p>In lecture 2 we have presented the graphical interface offered by MS Access
to work with base tables and virtual tables (views = select queries).  It is the 
simplest kind of an interface. It is called <i>the first level</i> (<i>elementary</i>). 
This lecture is devoted to the interface based on a more complex graphical object called
<i>a form</i>.  In the next lecture we will talk about <i>reports</i> and <i>web pages</i>.
All of them, i.e. forms, reports and web pages constitute the <i>intermediate level</i>.
The third, so called <i>professional level</i> is reached when the user's actions are
supported by code which automatically performs some functions like opening another form
with  the requested data or the preparation of data for a particular report.
Complete database applications are on the professional level.  It will be discussed further
during following lectures.</p> 

<p>Forms are used to enter data into a database and to present the data to the user.
Forms belong to the user's point of view.  A form can contain subforms.
Each part of a form (the main form and the subforms) is based on a table or a query
which
has been indicated as the data source for the part of the form.
The definition of a form is developed and presented in <i>the design view</i>.
The content of the form is shown in <i>the form view</i>.

<p><a href="#Zadania">The exercises</a> are integral part of this lecture.  If you solve them,
you will get a simple application for a library.</p>

<hr><h3><a name="Aplikacja">Database application</a></h3>

<p>A database application provides data stored in a database by means
of a graphical user interface. The fundamental task of the application is
<i>to respond to events</i> initiated by users or being the outcomes of certain actions
inside the application.  Such applications are called <i>event-driven</i>.</p>

<p>The application defines a set of scenarios which consists of steps. At each step the user
is presented a set of actions to choose from. The user need not be a programmer and most
probably he/she does not know the standard interface provided by the tools
used to build the application.</p>

<p>The basic item of the user interface is <i>a form</i> displayed on the screen.
A form consists of <i>controls</i>: text boxes, combo boxes, list boxes, labels,
pictures, charts, command buttons and ActiveX controls.  A form is used not only to
show the data but also to modify it (to insert new data and to delete or update 
existing data).</p>

<p>The data presented by a form is taken from a database which is managed by
a dedicated program - <i>the database server</i>. The database server may work
on the same computer as the application or be installed remotely.  The database server
may be run by the same software as the user interface (e.g. MS Access) or some other software
(e.g. SQL Server).  For the purpose of the present lectures 
we assume that both the server side
and the client side are run by MS Access. In the simplest case the objects of the user interface
are stored in the same database as the data.  However in real applications it is recommended to
separate them.</p>

<p>To sum up, the database application performs the following functions:

<ol>
<li>storage of large amounts of data and fast access to it
	(<i>the server side</i>),
<li>the graphical user interface to the database
	(<i>the client side</i>),
<li>insertions of new data into the database,
<li>updates and deletions of data in the database,
<li>searching the data in the database,
<li>the presentation of the found data by means of forms, reports and charts,
<li>processing of business data.
</ol>

<p><a name="MSAccess">MS Access</a> is a simple program which can be used to create
simple databases and database applications.  Here are the facilities of MS Access:</p>

<ol>
<li>object-based event-driven programming in VBA (= Visual Basic for Applications),
<li>facilities to create web interface (however, they work only with Internet Explorer),
<li>connectivity to other SQL databases,
<li>integration with MS Office.
</ol>

<h4><a name="Rodzaje">Types of objects in MS Access applications</a></h4>

<ul>
<li>a table (a database object),
<li>a query (a database object),
<li>a form,
<li>a report,
<li>a web page,
<li>a macro,
<li>a module (a collection of procedures and functions).
</ul>

<p><a name="Grupy">You can also divide objects of an application into <i>groups</i>
which are logical
chunks, e.g. the group of objects concerned with the customer service or
the group of objects concerned with the stock.  In order to create a group, click
the button
"Groups" and choose "New Group" from the menu.</a></p>

<hr><h3><a name="Formularz">Forms</a></h3>

<p>A form is the fundamental item in the graphical user interface of a database application.
An application includes a set of inter-connected forms.</p>

<p>A form can be used to:</p>

<ol>
<li>enter data into the database,
<li>present data to the user,
<li>modify data of the database,
<li>remove data from the database,
<li>print documents with data,
<li>initiate business actions in the information system.
</ol>

<p>Here are the basic properties of forms.</p>

<ol>
<li>A form represents the point of view of the user.
<li>A form may contain subobjects like charts and subforms.
<li>Each part of a form (the main form and the subforms) is based on a table or a query
	which 
	has been indicated as the data source for the part of the form.
<li>The definition of a form is developed and presented in <i>the design view</i>.
<li>The content of the form may be shown in three different views described below.
</ol>

<h4><a name="Pojed">Single Form</a></h4>

<p><i>Single Form</i> displays one record on the screen.  The fields
are put into one
column by default.</p>

<p align="center"><img border="0" src="https://gakko.pjwstk.edu.pl/materialy/2398/lec/wyklad06/images/6_6.png"></p>

<h4><a name="Ark">Datasheet</a></h4>

<p><i>Datasheet</i> displays a simple table like the datasheet for a table or a query.</p>

<p align="center"><img border="0" src="https://gakko.pjwstk.edu.pl/materialy/2398/lec/wyklad06/images/6_2.png"></p>

<h4><a name="Ci">Continuous Forms</a></h4>

<p><i>Continuous Forms</i> are either <i>columnar</i> or <i>tabular</i>.</p>

<h5>Columnar form</h5>

<p><i>Columnar form</i> displays the sequence of records. Each record is a column of fields
like on a Single Form.</p>

<p align="center"><img border="0" src="https://gakko.pjwstk.edu.pl/materialy/2398/lec/wyklad06/images/6_3.png"></p>

<h5>Tabular form</h5>

<p><i>Tabular form</i> displays one record in a row.</p>

<p align="center"><img border="0" src="https://gakko.pjwstk.edu.pl/materialy/2398/lec/wyklad06/images/6_4.png"></p>


<hr><h3><a name="Widok">Design view</a></h3>

<p>In <i>Design view</i> we create and place all controls of the form, e.g. text boxes
connected with columns of the database, text labels (constant strings).</p>

<p align="center"><img border="0" src="https://gakko.pjwstk.edu.pl/materialy/2398/lec/wyklad06/images/6_5.png"></p>
<p align="center"><img border="0" src="https://gakko.pjwstk.edu.pl/materialy/2398/lec/wyklad06/images/6_1.png"></p>

<hr><h3><a name="Pod">Subform</a></h3>

<p>Forms can be nested. A form can be in control of another form.  In such cases the nested
form is called <i>a subform</i>. A form with a subform usually renders data
from two tables connected with a many-one relationship.</p>

<p align="center"><img border="0" src="https://gakko.pjwstk.edu.pl/materialy/2398/lec/wyklad06/images/n_1.png"></p>

<p>And the same form in the design view:</p>

<p align="center"><img border="0" src="https://gakko.pjwstk.edu.pl/materialy/2398/lec/wyklad06/images/n_2.png"></p>

<p>A subform is a form itself.</p>

<hr><h3><a name="Stro">Pages</a></h3>

<p>If the amount of information in a single record is bigger than the screen,
the content of the form can be divided into <i>pages</i>. 
In order to change the current page, the user presses keys 
<b>PageUp</b> or <b>PageDown</b>.</p>

<p>If you want to use pages, set property <i>Cycle</i> of the form to 
<i>Current Page</i> and <i>Scroll Bars</i> to either <i>None</i> or
<i>Horizontal Only</i>.</p>

<p>The pages are particularly useful when you print the content of the form.</p>

<hr><h3><a name="Zak">Tabs</a></h3>

<p><i>Tabs</i> are also useful, e.g. to divide the view into basic information 
and Curriculum vitae (a field of type <i>Memo</i>).</p>

<p align="center"><img border="0" src="https://gakko.pjwstk.edu.pl/materialy/2398/lec/wyklad06/images/6_9.png"></p>

<hr><h3><a name="Panel">Control panel</a></h3>

<p><i>A control panel</i> is a special kind of form.  It shows no data.
It contains only buttons, labels and images.
A user may click one of them to perform an operation on
data. Such a form is not connected to any table or query.</p>

<p align="center"><img border="0" src="https://gakko.pjwstk.edu.pl/materialy/2398/lec/wyklad06/images/6_10.png"></p>

<hr><h3><a name="Kreator">Form wizard</a></h3>

<p><i>The form wizard</i> can create forms.
It asks the user to answer some questions and tries to build a form
which matches the requests of the user. If two tables are connected with 
a many-one relationship, we will have to choose one from many
possibilites.  For example let us consider tables <i>Persons</i>
and <i>Departments</i>.</p>

<ol>
<li>A form with a subform.  We choose table <i>Departments</i> for
	<i>the main form</i> and <i>Persons</i> for <i>the subform</i>.

	<p align="center"><img border="0" src="https://gakko.pjwstk.edu.pl/materialy/2398/lec/wyklad06/images/6_12.png"></p>
	
	<p>Here is the form created by the wizard:</p>

	<p align="center"><img border="0" src="https://gakko.pjwstk.edu.pl/materialy/2398/lec/wyklad06/images/6_14.png"></p>

	<p>This form can be used to view departments and their persons as well
	as to enter new departments and persons. You can also update and delete records
	unless you break the referential integrity.</p>

<li>A form with "Look up".  We choose table <i>Persons</i> and ask form additional 
	information on the department of each person.  This form will be based on a query
	which joins tables <i>Persons</i> and <i>Departments</i>.

	<p align="center"><img border="0" src="https://gakko.pjwstk.edu.pl/materialy/2398/lec/wyklad06/images/6_11.png"></p>

	<p>Here is the form created by the wizard:</p>

	<p align="center"><img border="0" src="https://gakko.pjwstk.edu.pl/materialy/2398/lec/wyklad06/images/6_13.png"></p>

	<p>This form can be used to view the data on each employee with the information
	on his/her department, to enter a new person and to assign him/her to an existing department.
	You can also update and delete records
	unless you break the referential integrity.</p>
</ol>



<hr><h3><a name="Arkusz">Properties of objects</a></h3>

<p>Properties of the form and its elements can be viewed and set in the property window.
To open this window, switch to Design view and choose "Properties" 
from the
pop-up menu, the toolbar or from menu "View". The property window of a form may look as follows:
</p>

<p align="center"><img border="0" src="https://gakko.pjwstk.edu.pl/materialy/2398/lec/wyklad06/images/n_3.png"></p>


<p><i>Record Source</i> is the fundamental property of a form.  It tells you where to
take records from.  The value of this property can be the name of a query or table.
In this example the record source is defined by the SELECT statement, which is not stored
on a separate name.</p>

<p>Note that the name of the form is not among those properties.
However, among the properties of a form there is 
<i>Caption</i>. As we are going to see, the name of the form is an attribute of the
object available in VBA code. However, the names of the elements of the form are available
as their property  <i>Name</i>.</p>

<hr><h3><a name="Switchboard">Switchboard Manager</a></h3>

<p><i>The switchboard manager</i> is a special utility program which
can be used
to create form menus. It is available as the menu item &quot;Tools
-&gt; Database Utilities -&gt; Switchboard Manager&quot;.</p>

<p align="center"><img border="0" src="https://gakko.pjwstk.edu.pl/materialy/2398/lec/wyklad06/images/sm2.png"></p>

<p>Here is the result:</p>

<p align="center"><img border="0" src="https://gakko.pjwstk.edu.pl/materialy/2398/lec/wyklad06/images/sm1.png"></p>


<hr><h3><a name="Czci">Sections of a form</a></h3>

<p>A form consists of the following <i>sections</i>:</p>

<ul>
<li><i>the form header</i> is usually the place of the information and the buttons
	which concern the form as a whole,
<li><i>the detail</i> is used to present the records (the detailed data),
<li><i>the form footer</i> is used to present the summary information calculated
	from the detailed data.
</ul>

<p>As <a=href="#Stro">we said before</a>, you can divide the form into pages.  You can
define a separate header and footer for each page.</p>

<hr><h3><a name="Dozwolone">Available operations on data</a></h3>

<p>The set of operations available through the form is defined by
the following properties than can be either "Yes" or "No":</p>


<ol>
<li><i>Allow edits</i>
<li><i>Allow deletions</i>
<li><i>Allow additions</i>
<li><i>Data entry</i> (it is only allowed to enter new data)
</ol>

<p>These properties can be used to state that the form may be used:</p>

<ol>
<li>only to enter new records (edits,deletions=<i>No</i>, additions,entry-only=<i>Yes</i>),
<li>only to read data from the database (4 times <i>No</i>),
<li>only to update the data without entering new data 
	(edits,deletions=<i>Yes</i>, additions,entry-only=<i>No</i>),
<li>to add new records and delete or update existing records
	(this is the default setting: 
	edits,deletions,additions=<i>Yes</i>, entry-only=<i>No</i>)
</ol>

<p>These properties allow setting the operations which can be performed by users through
each form. This way we can eliminate wrong usages of the form.  This facilitates 
preservation of the integrity.</p>

<hr><h3><a name="Elementy">Controls</a></h4>

<p>A form contains <i>controls</i> of the following kinds:</p>

<ol>
<li><i>bound</i> - the data source of such controls is a field of a table or a query, e.g.
	a text field. The easiest way to create a text box is to use the <i>Field List</i>
	(choose menu item &quot;View -&gt; Field List&quot; or tool &quot;Field List&quot;).
	Drag the selected field and drop it somewhere on the form.
<li><i>unbound</i> - e.g. a field with a constant value, label, line and image.
<li><i>derived</i> - the data source for it is an expression, e.g. 
	<code>= [Net price]*1,22</code>
</ol>

<p>If a name contains special characters (e.g. spaces), you have to surround it with brackets.
Of course, you can use brackets with any name,
even though you do not need to.</p>

<hr><h3><a name="Wyrazenia">Expressions</a></h3>

<p><i>Expression</i> allow transforming data retrieved from the database
to the form convenient for the user.</p>

<p>If you define the expression for a derived field,
precede it with the equality sign.
The arguments of functions are separated by:

<ul>
<li>commas (in SQL statements and VBA code) or
<li>semicolons (in the design view),
</ul>

<h4><a name="Expression">Expression Builder</a></h4>

<p><i>Expression Builder</i> is a useful tool which helps writing 
expressions, e.g. the values of fields' properties &quot;Control Source&quot;
and &quot;Default Value&quot;.  This tool is fired when you click the
button with dots
or press SHIFT+F2.</p>

<p align="center"><img border="0" src="https://gakko.pjwstk.edu.pl/materialy/2398/lec/wyklad06/images/7_1.png"></p>

<h4><a name="Przyk">Examples</a></h4>

<table border="1" align="center">

<tr>	<td width="40%"><code>[Net price]*0,75<br> 
		[Partial sum]+[Commission]</code>
	<td>arithmetic operations

<tr>	<td><code>[City]&amp;&quot; &quot;&amp;[Zip]</code>
	<td>concatenation of strings
<tr>	<td><code>DateSerial(2002,9,4)</code>
	<td>date September 4, 2002
<tr>	<td><code>Date()</code>
	<td>current date
<tr>	<td><code>Now()</code>
	<td>current date and time
<tr>	<td><code>Sum([Commission])<br>
    		Count([Name)<br>
		Max([Sal]), Min([Sal])<br>
		Avg([Sal])</code>
	<td>aggregates in the footer of a form or a report
<tr>	<td><code>&quot;Page &quot;&amp;[Page]&amp;&quot; of &quot; &amp;[Pages]</code>
	<td>current page and number of pages in the footer of a form or a report
<tr>	<td><code>IIf(IsNull([Sal]),0,[Sal])</code>
	<td>interpret <code>Null</code> as zero
<tr>	<td><code>Left([Region],1)<br>
		Right([Region],1)</code>
	<td>first letter of values of <code>[Region]</code><br>
	last letter of values of <code>[Region]</code><br>
<tr>	<td><code>Middle([Phone no],2,3)</code>
	<td>three characters starting from the second character of
		<code>[Phone no]</code>
<tr>	<td><code>DatePart(format, date)<br>
    	DatePart(&quot;yyyy&quot;, [Hired])</code>
	<td>fragment of the date; format <code>yyyy</code>
		means the four digits of the year
<tr>	<td><code>DateAdd(&quot;d&quot;, -10, [Promised])<br>
    	[Promised]-10</code>
	<td>ten days before date <code>[Promised]</code>

<tr>	<td><code>DateDiff(&quot;d&quot;, [Ordered], [Shipped])<br>
	[Shipped]-[Ordered]</code>
	<td>the number of days between dates <code>[Ordered]</code>
	and <code>[Shipped]</code>
<tr>	<td><code>[Author] Like &quot;Lech*&quot;</code>
	<td> true if the string matches the pattern; wildcard characters are
	<code>*</code> = any number (even zero) of characters and
	<code>?</code> = any character
<tr>	<td><code>[Price] BETWEEN 1000 AND 2000</code>
	<td>equivalent to <code>1000 &lt;= [Price] AND [Price] &lt;= 2000</code>
</table>

<h4><a name="Odwolania">Referencing controls</a></h4>

<p>In expressions we can reference the controls of forms and reports.  These
references have the following form:</p>

<pre><b>Forms</b>![<i>name_of_form</i>]![<i>name_of_control</i>]
<b>Reports</b>![<i>name_of_report</i>]![<i>name_of_control</i>]</pre>

<p>The referenced form or report must be opened.</p>

<h5>Examples</h5>

<p>The value of expression:

<blockquote><code>Forms![Persons]![Name]</code></blockquote>

is the text entered into the field <i>Name</i> from the 
form <i>Persons</i> which must be opened.
The same expression can be used to set the value of this field:

<blockquote><code>Forms![Persons]![Name] = "Smith"</code></blockquote>

Use the dot to reference the properties of a form, a report or a control:

<pre><b>Forms</b>![<i>name_of_form</i>].[<i>property</i>]
<b>Forms</b>![<i>name_of_form</i>]![<i>control</i>].[<i>property</i>]</pre>

<h5>Example</h5>

<blockquote><code>Forms![Persons].[Record Source]<br>
Forms![Persons]![Gender].[Default Value]</code></blockquote>

You can also change the value of a property:

<blockquote><code>Forms![Persons]![Gender].[Default Value] = 'Female'</code></blockquote>


<hr><h3><a name="Zestaw">Toolbox</a></h3>

<p>Controls are put onto the form by means of the toolbox. You can open it with
menu item &quot;View -&gt; Toolbox&quot;.</p>

<table border="1">
<tr><td valign="center">
	<ol>
	<li>Select objects
	<li>Label
	<li>Option group
	<li>Option button
	<li>Combo box
	<li>Command button
	<li>Unbound object frame
	<li>Page break
	<li>Subform/subreport
	<li>Rectangle
	</ol>
<td><img border="0" src="https://gakko.pjwstk.edu.pl/materialy/2398/lec/wyklad06/images/6_15.png"/>
<td valign="center">
	<ol>
	<li>Control wizards
	<li>Text box
	<li>Toggle button
	<li>Check box
	<li>List box
	<li>Image
	<li>Bound object frame
	<li>Tab control
	<li>Line
	<li>More controls
	</ol>
</table>

<p>You can create controls with or without the wizard.  To toggle the wizard
use the button "Control wizards".  Here is more information on these controls:

<ol>
<li>There are two kinds of lists.
	<ol type="a">
	<li><i>List box</i> shows the allowable values on a vertical list.
	<li><i>Combo box</i> allows entering text and offers a pull-down list of choices.
	</ol>
<li><i>Option buttons</i>, <i>check boxes</i> and <i>toggle buttons</i> are used to set and to show binary
	values (<i>Yes</i> = -1 and <i>No</i> = 0). 
<li><i>Option group</i> consists of the group frame and a set of its option buttons,
	check boxes and toggle buttons. The opting group is bound to a number field.
	Subsequent integers (1, 2, 3, ...) represent the available options.
<li><i>OLE frame</i> may be unbound (it shows a constant image) or bound
	(shows OLE objects stored in the database). OLE objects are
	MS Word documents, Excel spreadsheets.  In order to add a new OLE object
	choose the menu item &quot;Insert -&gt; Object&quot;.
<li><i>Charts</i> are created by wizards.
<li><i>Subform</i> shows a source form which was created before.
	It can be synchronized with the main form by means of common fields
	(if you use the wizard, it will create the link for you automatically).
<li><i>Command buttons</i> are associated with a macro or a procedure.
</ol>

<p align="center"><img border="0" src="https://gakko.pjwstk.edu.pl/materialy/2398/lec/wyklad06/images/6_16.png"></p>

<hr><h3><a name="Pola">List box</a></h3>

<p>List boxes are very useful controls which can implement various features. A list box
can contain:</p>

<ul>
<li><i>a fixed list of values</i>, e.g. the names of weekdays, the names of months;
<li><i>a dynamic list of values</i> read from the database, e.g. the allowable values
	for the column of a foreign key taken from the appropriate column of the
	primary key of the associated table.
</ul>

<h5>Example</h5>

<p>The list box can be used to display allowable values for foreign keys:

<ul>
<li>the list labeled <i>Customer</i> shows the names of the customers of the company.
<li>the list labeled <i>Responsible</i> shows the names of the employees of the company.
</ul>

<p align="center"><img border="0" src="https://gakko.pjwstk.edu.pl/materialy/2398/lec/wyklad06/images/6_17.png"></p>


<p>You have to distinguish between two properties of the text box.</p>

<ul>
<li><i>Row Source</i> is the source of the data displayed by the text box.
<li><i>Control Source</i> is the destinations of the chosen value. If it is set
	we call the text box <i>bounded</i>.  Otherwise it is <i>unbounded</i>.
</ul>


<h4><a name="Kreat">Combo box wizard</a></h4>


<p>The wizard can create the combo box which is used to find a particular record.
When the user selects data from this combo box, MS Access displays the record
which contains this data.  In order to create such a combo box, select the third option
offered by the wizard.</p>

<p align="center"><img border="0" src="https://gakko.pjwstk.edu.pl/materialy/2398/lec/wyklad06/images/6_19.png"></p>

<p>Choose <i>Name</i> of <i>Employee</i> as the <i>Row Source</i>.  This will cause an
unbounded combo box to be created.  It will display the list of employees.  When you choose
the name of an employee from the pull-down list, the data on this employee will be displayed
in the detail section.</p>

<p align="center"><img border="0" src="https://gakko.pjwstk.edu.pl/materialy/2398/lec/wyklad06/images/6_20.png"></p>


<h4><a name="Przycisk">Command button wizard</a></h4>

<p>The command button wizard is fired when you drag 
the command button from the toolbox 
to the form.  In the window which appears you should
choose the command to be executed
when a user clicks the button, e.g. closing the form.</p>

<p align="center"><img border="0" src="https://gakko.pjwstk.edu.pl/materialy/2398/lec/wyklad06/images/b_1.png"></p>


<p>Here is the result:</p>

<p align="center"><img border="0" src="https://gakko.pjwstk.edu.pl/materialy/2398/lec/wyklad06/images/b_2.png"></p>

<p align="center"><img border="0" src="https://gakko.pjwstk.edu.pl/materialy/2398/lec/wyklad06/images/b_3.png"></p>

<hr><h3><a name="Wyszukiwanie">Searching the data</a></h3>

<p>By means of built-in tools and menus &quot;Edit&quot; and &quot;Records&quot;
you can filter the set of displayed records or search a particular record
(in tables, queries and forms).  You will also find appropriate buttons on the toolbar:</p>

<table align="center">
<tr><td colspan=6><img border="0" src="https://gakko.pjwstk.edu.pl/materialy/2398/lec/wyklad06/images/6_18.png">
<tr>	<td width="16%">1
	<td width="16%">1
	<td width="16%">2
	<td width="16%">3
	<td width="16%">5
	<td width="16%">6
</table>

<p>Here is the list and the explanation of the functions of these buttons:</p>


<ol>
<li><i>Sort ascending</i>, <i>Sort descending</i> - order by the current field.
<li><i>Filter by Selection</i> - filter out records which have other value in the 
	currently selected field. 
<li><i>Filter by form</i> - filter out records which have other value in the 
	currently selected field. 
<li><i>Advanced filter/sort</i> - available only from menu "Records"; display
	the query spreadsheet to enter the filtering condition. Then apply this condition.
<li><i>Apply filter/sort</i> - display a form to enter the filtering condition.
	Then apply this condition.
	<ul>
	<li>To cancel the filter, click the same button which is currently named &quot;Remove filter&quot;.
	<li>Most recently applied filter is stored in property <i>Filter</i> of the form.
	<li>Most recently applied order is stored in property <i>Order by</i> of the form.
	</ul>
<li><i>Find</i> - find records with the specified text.  Available options are:
	"Any Part of Field", "Start of Field", "Whole Field", "Up", "Down", "Match case",
	"Search Fields as Formatted". 	
</ol>


<hr><h3><a name="Podsumowanie">Summary</a></h3>

<p>Every database application consists of at least two parts:

<ol>
<li><i>the database</i> (<i>the server side</i>)
	which stores large amounts of data and provides fast access
	to that data and
<li><i>the graphical user interface</i> = <i>GUI</i> (<i>the client side</i>)
	which provides functions for the users. 
	These functions include operations concerned with storing data and
	performing actions on the stored data.
</ol> 

<p>Forms are the fundamental elements  of the graphical user interface.
They are used to enter data into a database and to present that data to the user.
They belong to the user's view.  A form can contain subforms.
Each part of a form (the main form and the subforms) is based on a table or a query
which
has been indicated as the data source for the part of the form.
The definition of a form is developed and presented in <i>the design view</i>.
The content of the form is shown in <i>the form view</i>.</p>


<hr><h3><a name="Slownik">Dictionary</a></h3>

<dl>
<dt><a href="#Elementy">control</a>
<dd>An item which facilitates interaction with the user, e.g. a text box, a list box, a button,
	an image, a message window.
<dt><a href="#Panel">control panel</a>
<dd>A kind of form which displays buttons that start operations on data. 
	If the user wants to perform an operation, he/she clicks the appropriate button.
<dt><a href="#Aplikacja">database application</a>
<dd>An application with a graphical user interface which provides 
	the access to the data stored in a database.
<dt><a href="#Formularz">form</a>
<dd>the fundamental element of the graphical user interface in a database application.  A database application
	contains a set of inter-connected forms.
<dt><a href="#Stro">page</a>
<dd>A way to divide the form into parts.
<dt><a href="#Czci">section</a>
<dd>A part of the form, either the header, the detail or the footer.

<dt><a href="#Zak">tab</a>
<dd>A way to divide the form into parts.

<dt><a href="#Zestaw">toolbox</a>
<dd>A collection of icons which are used to create controls on a form.
</dl>

<hr><h3><a name="Zadania">Exercises</a></h3>

<h4>Information</h4>

<p>In order to make the following exercises, you need database <i>Library</i> that is the result
of exercises from lecture 3.</p>

<h4><a name="Zadanie 1">Exercise 1</a></h4>

<p>Make the following changes in this database:</p>

<ol>
<li>Convert the columns of foreign keys into lookups.
<li>Add the following fields to table <i>Books</i>
	<ol type="a">
	<li><i>Textbook</i> (<i>Yes</i>/<i>No</i>);
	<li><i>Medium</i> of number type <i>Byte</i> as the code of the medium: hardcover, paperback, 
		CD, microfilm; set 1 as the default value.
	</ol>
<li>Enter sample data into the tables.
</ol>

<h4>Exercise 2</h4>

<p>Create form <i>Books</i> without any help from the wizard.</p>

<ol>
<li>In the database window choose "Forms" and click "New". Select "Design View" and table <i>Books</i> 
	from the combo box.
<li>Activate the toolbox and the field list (either from the toolbar or from menu "View").
<li>Select all fields on the field list (with SHIFT). Drag them onto the form. Switch to Form View
	(select appropriate item from the toolbar, pop-up menu or menu "View").
<li>Walk through all displayed records by means of the navigation button visible 
	at the bottom of the form. Insert a new record. Walk through all the records again.
	Find the new record and remove it. Examine the options available in
	the menu "Edit".
<li>Switch back to the design view. Open Properties by means of the toolbar, pop-up menu or menu "View").
	The title of the window should be "Form".  Select tab "Format" and set the "Caption" of your form.
	Check the look of your form through setting the "Default View" to "Continuous Forms" and then
	"Datasheet".  After that, set this property to "Single Form" again. 
<li>Before starting with this point make a copy of your form and employ your imagination
	on this copy.  Change the look and feel of the form and its control as you wish,
	Are you happy with this created from?
	<ol type="a">
	<li>Move text boxes and labels. Change their sizes. To select objects, use the mouse button.
		To select more than one, keep SHIFT key pressed. When you select a control, you can
		move it and change its size by means of arrows, SHIFT and CTRL.
		Examine various handles which appear when you move the mouse pointer over controls.
		Are you able to move the text box without its label?
		Are you able to delete the label only?
	<li>If you select a group of controls you can align their positions and sizes by means
		of the items "Align" and "Size" from the menu "Format".
	<li>Insert new text labels which display the functions of the form.
	<li>Choose appropriate image for the background of your form, i.e. set
		the property "Picture"
		of the form.  Try the item "Auto Format" from the menu "Format".
	<li>Do you like the look of all controls of the form?
		Visit the properties of each control and try to set appropriate values of
		properties like "Back Style", "Back Color", "Special Effect",
		"Border Style", "Border Color", "Border Width", "Font Color", "Font Name",
		"Font Size", "Font Weight", "Font Italic".  Remember that all these properties
		can also be set by means of the toolbar "Formatting".
	</ol>
<li>Convert the combo box with the domain/subject of the book to a list box (choose
	the menu item
	"Format -&gt;  "Change To").  Check the change of the appearance in the form view.
	Perhaps the combo box was better.
<li>Convert check box <i>Textbook</i> to a toggle button an then to an option button
	("Format -&gt;  "Change To").  Check the change of the appearance in the form view.
<li>Remove text box <i>Medium</i>.  Switch off the wizard in the toolbox. Create
	an option group
	<i>Medium</i>.  Switch to tab "Data" and set property "Control source" to <i>Medium</i>
	(select it from the drop-down list).  Choose the kind of items of this option group
	(a check box, an option button or a toggle button). Create one control for each
	type of media. Set appropriate text labels from the option group and its elements.
<li>Create a select query which returns the fields: <i>ISBN</i> of table <i>Authors</i>
	and <i>First Name</i> and <i>Last Name</i> from table <i>Persons</i>.

<li>Create a continuous form <i>Authors of books</i> with the data source
	being the query created in the previous point. If you can resist
	using the wizard, choose the option "Autoform: Tabular" in the window "New
	Form". Hide the field <i>ISBN</i> (i.e. sets its property <i>Visible</i>
	to <i>No</i>) and move it to the end.  Minimize its size and remove
	its label.

<li>Return to the form <i>Books</i>. Choose the control <i>Subform/subreport</i>
	from the toolbox and place it at the bottom of the form <i>Books</i>. 
	Set the property <i>Source object</i> of the new subform to the name of
	the form created in the previous point (<i>Authors of books</i>).  You
	can also select it from the drop-down list. Check whether the properties
	<i>Link child fields</i> and <i>Link master fields</i> are set to
	<i>ISBN</i>. Switch to the form view. Usually the content of the
	subform is not presented properly at once. Switch to the design view
	and correct the size of the subform and the form linked as the
	subform itself.

<li>Use the same procedure to build subforms which display translators of
	each book and its domain/subject. Use this form to display all the
	books from the database.  Congratulate yourself. Now you are able to
	design nontrivial graphical user interfaces to the database.

</ol>

<h4>Exercise 3</h4>

<p>Create a form <i>Persons</i> which displays complete information on persons.  The data
on each person should be accompanied with the list of books authored or translated
by this person and the list of books borrowed from the library by this person.
</p>
