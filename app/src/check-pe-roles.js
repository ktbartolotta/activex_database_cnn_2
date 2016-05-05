
var cnnStr = (function () {
  return {
    PEPRD: "Provider=msdaora;Data Source=host1;User id=user;Password=pass1",
    PADH:  "Provider=msdaora;Data Source=host2;User id=user;Password=pass2"
  }
}());

$(function () {
  var persons = getPersonNames();
  $("#person").autocomplete({
    source: persons,
    minLength: 5
  });

  $("#launchButton").click(function () {
    var person = $("#person").val();
    if (person != '') {
      var id = getEmployeeId(person);
      $("#personId").text(id);
      setPEWorkAssignmentsById(id);
    }
  });
});

function getPersonNames() {
  var query = "select name from ppltempls";
  var cnn = new Connection(cnnStr.PADH);
  var persons = new Array();

  cnn.open();
  var rs = cnn.execute(query);
  if (!rs.EOF) {
    rs.MoveFirst();
    while (!rs.EOF) {
      var str = '';
      str += rs.fields("name");
      persons.push(str);
      rs.MoveNext();
    }
  }

  rs.Close();
  cnn.close();
  persons.sort();
  return persons;
}

function getEmployeeId(name) {
  var query = "select oprid from ppltempls where name = '" + name + "'";
  var cnn = new Connection(cnnStr.PADH);
  var id = '';

  cnn.open();
  var rs = cnn.execute(query);
  if (!rs.EOF) {
    rs.MoveFirst();
    id += rs.fields("oprid");
  }

  rs.Close();
  cnn.close();
  return id;
}

function setPEWorkAssignmentsById(id) {
  var assignments = new Array();

  var query = "select bmi_business_unit_db_id, pe_work_assignments_db_id, "
    + "lead_flag from pe_person_work_assignments where active_flag = 'Y' and "
    + "hrpr_per_db_id = '" + id + "' "
    + "order by bmi_business_unit_db_id, pe_work_assignments_db_id";
  var cnn = new Connection(cnnStr.PEPRD);
  var roleTable$ = $("#roleTable");

  roleTable$.html("<tr><th>Business Unit</th><th>Work Assignment</th><th>Lead Flag</th></tr>");
  cnn.open();
  var rs = cnn.execute(query);
  if (!rs.EOF) {
    rs.MoveFirst();
    while (!rs.EOF) {
      roleTable$.append(
	  "<tr><td>" + getBMIBusinessUnitById(rs.fields("bmi_business_unit_db_id")) + "</td>"
	  + "<td>" + getWorkAssignmentById(rs.fields("pe_work_assignments_db_id")) + "</td>"
	  + "<td>" + rs.fields("lead_flag")  + "</td></tr>");
      rs.MoveNext();
    }
  }

  rs.Close();
  cnn.close();
  return assignments;
}

function getBMIBusinessUnitById(id) {
  var query = "select * from pe_ppl_business_units where db_id = " + id;
  var cnn = new Connection(cnnStr.PEPRD);
  var businessUnit = '';

  cnn.open();
  var rs = cnn.execute(query);
  if (!rs.EOF) {
    rs.MoveFirst();
    businessUnit += rs.fields("business_unit_descr");
  }

  rs.Close();
  cnn.close();
  return businessUnit;
}

function getWorkAssignmentById(id) {
  var query = "select * from pe_work_assignments where db_id = " + id;
  var cnn = new Connection(cnnStr.PEPRD);
  var assignment = '';

  cnn.open();
  var rs = cnn.execute(query);
  if (!rs.EOF) {
    rs.MoveFirst();
    assignment += rs.fields("name");
  }

  rs.Close();
  cnn.close();
  return assignment;

}

function Connection(connectionString) {
    this.cnn = new ActiveXObject('ADODB.Connection');
    this.cnnStr = connectionString;

    this.open = function () {
        this.cnn.Open(this.cnnStr);
    }

    this.close = function () {
        this.cnn.Close();
    }

    this.execute = function (query) {
        var rs = this.cnn.Execute(query);
        return rs;
    }
}
