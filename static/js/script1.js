var frame_count = 1;
function add_row(elem){
	
	var row_index = elem.parentElement.parentElement.children[0].rows.length - 1;
	var row = elem.parentElement.parentElement.children[0].rows[row_index];
    var table = elem.parentElement.parentElement.children[0];
    var clone = row.cloneNode(true);
	var cells_length = clone.cells.length - 1;
	for(var i = 1; i<cells_length; i++){
		clone.cells[i].children[0].value=""
	}
    table.appendChild(clone);
}

function delete_row(elem){
	if(elem.parentElement.parentElement.parentElement.rows.length > 2){
		elem.parentElement.parentElement.remove();
	}else{
		alert("First row can not be deleted!")
	}
}


function submit(){
	
	
	var frame_list = document.getElementsByClassName("frame");
	var basic_details = {};
	if ($('#company_name').val()=="")
	{
		alert("Please select company Name");
		return;
	}
	
	if ($('#ahu_number').val()=="")
	{
		alert("Please Enter AHU Number");
		return ;
	}
	if ($('#room_name').val()=="")
	{
		alert("Please Room Name");
		return ;
	}
	if ($('#room_volume').val()=="")
	{
		alert("Please enter Room Volume");
		return ;
	}	
	if ($('#grade').val()=="")
	{
		alert("Please select Grade");
		return ;
	}	
	if ($('#acph_thresold').val()=="")
	{
		alert("Please enter ACPH Thresold");
		return ;
	}	
	if ($('#SR_NO').val()=="")
	{
		alert("Please select Serail Number");
		return ;
	}
	if ($('#location').val()=="")
	{
		alert("Please enter Department ");
		return ;
	}
	if ($('#Test_taken').val()=="")
	{
		alert("Please Enter Test_taken Date");
		return ;
	}
	var MAKE_MODEL = $('#MAKE_MODEL').val()
	var full_data = {};
	var final_table_data= {};
	basic_details['sr_no'] =$('#SR_NO').val();
	basic_details['company_name'] =$('#company_name').val();
	basic_details['room_volume'] =$('#room_volume').val();
	basic_details['room_name'] =$('#room_name').val();
	basic_details['ahu_number'] =$('#ahu_number').val();	
	basic_details['location'] =$('#location').val();
	basic_details['Test_taken'] =$('#Test_taken').val();
	basic_details['grade'] =$('#grade').val();
	basic_details['acph_thresold'] =$('#acph_thresold').val();
		
	code_tbl = frame_list[0].getElementsByClassName("code_tbl")[0]
	code_rows = code_tbl.rows
	for(var j = 1; j<code_rows.length; j++)
	{
		tds = code_rows[j].children
		
		if (tds[0].firstElementChild.value=="")
		{	
				alert("Filter ID number cannot be blank in row : "+j);
				return;
		}
		
		for(var l = 1; l<6; l++)
		{

			if (tds[l].firstElementChild.value=="")
			{	
				alert("Please enter correct value for V"+l+" in row : "+j);
				return;
			}
			if (MAKE_MODEL !="HSTENI")
			{
				if (tds[l].firstElementChild.value % 10!=0)
				{
					alert("Please enter number divisble by 10 for V"+l+" in row : "+j);
					return;
				}
			}
		}
		
		if (tds[6].firstElementChild.value=="")
		{	
				alert("Inlet Size cannot be blank in row : "+j);
				return;
		}
		
		
		
		
		
	}
	for(var j = 1; j<code_rows.length; j++){
		
		tds = code_rows[j].children
		Label_number = tds[0].firstElementChild.value
		V1 = tds[1].firstElementChild.value
		V2 = tds[2].firstElementChild.value
		V3 = tds[3].firstElementChild.value
		V4 = tds[4].firstElementChild.value
		V5 = tds[5].firstElementChild.value
		inlet_size = tds[6].firstElementChild.value
		
		var table_data = {};
		table_data['Label_number'] =Label_number 
		table_data['V1'] =V1
		table_data['V2'] =V2
		table_data['V3'] =V3
		table_data['V4'] =V4
		table_data['V5'] =V5
		table_data['Inlet_size'] =inlet_size
		final_table_data[j] = table_data
	}
	
	full_data['observation']=final_table_data
	full_data['basic_details']=basic_details
	
	
	$.getJSON('/submit_data', 
	{
		params_data : JSON.stringify(full_data)
	}, function(result) 
	{
		var link = document.createElement('a')
		link.href =result.file_path;
		link.download = result.file_name;
		link.dispatchEvent(new MouseEvent('click'));
		

		
	});
	
}


function submit_consolidated(){
	
	if ($('#company_name').val()=="")
	{
		alert("Please select company Name");
		return;
	}
	
	if ($('#report_type').val()=="")
	{
		alert("Please Select Report Type");
		return ;
	}
	if ($('#start_date').val()=="")
	{
		alert("Please enter Start Date");
		return ;
	}
	if ($('#end_date').val()=="")
	{
		alert("Please enter end Date");
		return ;
	}
	var full_data = {};
	full_data['company_name'] = $('#company_name').val()
	full_data['report_type']  = $('#report_type').val()
	full_data['start_date']   = $('#start_date').val()
	full_data['end_date']     = $('#end_date').val()
	
	
	$.getJSON('/submit_consolidated', 
	{
		params_data : JSON.stringify(full_data)
	}, function(result) 
	{
		alert(result.msg);
		
	});
	
}






function submit_pao(){
	
	
	var frame_list = document.getElementsByClassName("frame");
	var basic_details = {};
	if ($('#company_name').val()=="")
	{
		alert("Please select company Name");
		return;
	}
	
	if ($('#ahu_number').val()=="")
	{
		alert("Please Enter AHU Number");
		return ;
	}
	if ($('#room_name').val()=="")
	{
		alert("Please Room Name");
		return ;
	}
	if ($('#SR_NO').val()=="")
	{
		alert("Please select Serail Number");
		return ;
	}
	if ($('#location').val()=="")
	{
		alert("Please enter Department ");
		return ;
	}
	if ($('#Test_taken').val()=="")
	{
		alert("Please Enter Test_taken Date");
		return ;
	}
	
	if ($('#compresed_value').val()=="")
	{
		alert("Please Enter Compressed Air Value");
		return ;
	}
	
	var MAKE_MODEL = $('#MAKE_MODEL').val()
	var full_data = {};
	var final_table_data= {};
	basic_details['sr_no'] =$('#SR_NO').val();
	basic_details['company_name'] =$('#company_name').val();
	basic_details['room_name'] =$('#room_name').val();
	basic_details['ahu_number'] =$('#ahu_number').val();	
	basic_details['location'] =$('#location').val();
	basic_details['Test_taken'] =$('#Test_taken').val();
	
	basic_details['compresed_value'] =$('#compresed_value').val();
	basic_details['check_val'] =$('#check_val').val();
	
		
	code_tbl = frame_list[0].getElementsByClassName("code_tbl")[0]
	code_rows = code_tbl.rows
	for(var j = 1; j<code_rows.length; j++)
	{
		tds = code_rows[j].children
		
		if (tds[0].firstElementChild.value=="")
		{	
				alert("Filter ID number cannot be blank in row : "+j);
				return;
		}
		
		if (tds[0].firstElementChild.value=="")
		{	
				alert("Filter INLET NUMBER cannot be blank in row : "+j);
				return;
		}
		
		
		if (tds[1].firstElementChild.value=="")
		{	
				alert("Filter Upstream cannot be blank in row : "+j);
				return;
		}
		
		
		
		if (tds[2].firstElementChild.value=="")
		{	
				alert("Filter Leakage cannot be blank in row : "+j);
				return;
		}
		
		if (tds[3].firstElementChild.value=="")
		{	
				alert("Filter Remark cannot be blank in row : "+j);
				return;
		}
		
		
		
		
		
	}
	for(var j = 1; j<code_rows.length; j++){
		
		tds = code_rows[j].children
		INLET_NUMBER = tds[0].firstElementChild.value
		Upstream = tds[1].firstElementChild.value
		Leakage = tds[2].firstElementChild.value
		Remark = tds[3].firstElementChild.value
		
		var table_data = {};
		table_data['INLET_NUMBER'] =INLET_NUMBER 
		table_data['Upstream'] =Upstream
		table_data['Leakage'] =Leakage
		table_data['Remark'] =Remark
		final_table_data[j] = table_data
	}
	
	full_data['observation']=final_table_data
	full_data['basic_details']=basic_details
	
	
	$.getJSON('/submit_data_pao', 
	{
		params_data : JSON.stringify(full_data)
	}, function(result) 
	{
		var link = document.createElement('a')
		link.href =result.file_path;
		link.download = result.file_name;
		link.dispatchEvent(new MouseEvent('click'));
		

		
	});
	
}




function submit_particle_report(){
	
	
	var frame_list = document.getElementsByClassName("frame");
	var basic_details = {};
	if ($('#company_name').val()=="")
	{
		alert("Please select company Name");
		return;
	}
	
	if ($('#ahu_number').val()=="")
	{
		alert("Please Enter AHU Number");
		return ;
	}
	if ($('#room_name').val()=="")
	{
		alert("Please Room Name");
		return ;
	}	
	if ($('#SR_NO').val()=="")
	{
		alert("Please select Serail Number");
		return ;
	}
	if ($('#location').val()=="")
	{
		alert("Please enter Department ");
		return ;
	}
	if ($('#grade').val()=="")
	{
		alert("Please Select Grade ");
		return ;
	}
	if ($('#condition').val()=="")
	{
		alert("Please Select condition");
		return ;
	}
	if ($('#Test_taken').val()=="")
	{
		alert("Please Enter Test_taken Date");
		return ;
	}
	
	if ($('#gl_value').val()=="")
	{
		alert("Please Select Guidelines");
		return ;
	}
	
	
	var MAKE_MODEL = $('#MAKE_MODEL').val()
	var full_data = {};
	var final_table_data= {};
	basic_details['sr_no'] =$('#SR_NO').val();
	basic_details['company_name'] =$('#company_name').val();
	basic_details['room_name'] =$('#room_name').val();
	basic_details['ahu_number'] =$('#ahu_number').val();	
	basic_details['location'] =$('#location').val();
	basic_details['Test_taken'] =$('#Test_taken').val();
	basic_details['condition'] =$('#condition').val();
	basic_details['grade'] =$('#grade').val();
	basic_details['gl_value'] =$('#gl_value').val();
	
	
		
	code_tbl = frame_list[0].getElementsByClassName("code_tbl")[0]
	code_rows = code_tbl.rows
	for(var j = 1; j<code_rows.length; j++)
	{
		tds = code_rows[j].children
		
		if (tds[0].firstElementChild.value=="")
		{	
				alert("Filter Location cannot be blank in row : "+j);
				return;
		}
		
		if (tds[1].firstElementChild.value=="")
		{	
				alert("Filter ≥ 0.5 μm cannot be blank in row : "+j);
				return;
		}
		
		
		if (tds[2].firstElementChild.value=="")
		{	
				alert("Filter ≥ 5.0 μm cannot be blank in row : "+j);
				return;
		}
		if (tds[3].firstElementChild.value=="")
		{	
				alert("Remark cannot be blank in row : "+j);
				return;
		}
		
		
		
		
	}
	for(var j = 1; j<code_rows.length; j++){
		
		tds = code_rows[j].children
		Location = tds[0].firstElementChild.value
		zeor_point_five = tds[1].firstElementChild.value
		five_point_zero = tds[2].firstElementChild.value
		remark = tds[3].firstElementChild.value
		
		
		var table_data = {};
		table_data['Location'] =Location 
		table_data['zeor_point_five'] =zeor_point_five
		table_data['five_point_zero'] =five_point_zero
		table_data['remark'] =remark
		final_table_data[j] = table_data
	}
	
	full_data['observation']=final_table_data
	full_data['basic_details']=basic_details
	
	
	$.getJSON('/submit_particle_report', 
	{
		params_data : JSON.stringify(full_data)
	}, function(result) 
	{
		var link = document.createElement('a')
		link.href =result.file_path;
		link.download = result.file_name;
		link.dispatchEvent(new MouseEvent('click'));
		

		
	});
	
}


















$(function() 
{
    $("#company_name").change(function() 
    {
	company_name = $('option:selected',this).text();
	$.getJSON('/update_company_details', 
	{
		params_data : JSON.stringify(company_name)
	}, function(data) 
	{
		$('#company_address').val(data.company_address);
      	$('#report_id').val(data.report_id);
        $('#Test_taken').val(data.test_taken);		
		$('#location').val(data.location);							
	});
 });
});

$(function() 
{
	$("#SR_NO").change(function()
	{        
	SR_NO = $('option:selected',this).text();
	$.getJSON('/update_instument_details', 
	{
		params_data : JSON.stringify(SR_NO)
	}, function(data) 
	{			
		$('#INSTRUMENT_NAME').val(data.INSTRUMENT_NAME);
      	$('#MAKE_MODEL').val(data.MAKE);
        $('#VALIDITY').val(data.VALIDITY);							
	
	});
  });
});


$('#paotable').on('change', 'input', function () {
    var row = $(this).closest('tr');
   	var check_val = $('#check_val').val();
    var upstream  = $('#Upstream', row).val();
	var Leakage   = $('#Leakage', row).val();	
	
	if ((upstream!="") && (upstream<=80 &&  upstream>=20) && (Leakage!="" )&& (Leakage<=check_val))
	{
		$('#Remark', row).val("Pass");
	}	
	else {$('#Remark', row).val("Fail");}
	
	
	
});

var val_1;
var val_2;
$("#grade").change(function()
	{
        
	grade = $('option:selected',this).text();
	var full_data = {};
	full_data['grade'] = grade
	full_data['gl_value'] = $('#gl_value').val()
	full_data['condition'] = $('#condition').val()
	$.getJSON('/get_limits', 
	{
		params_data : JSON.stringify(full_data)
	}, function(data) 
	{		
			val_1 = data.value1 
			val_2 = data.value2
										
	});
    });
	
$('#particleCountTable').on('change', 'input', function () {
    var row = $(this).closest('tr');
    console.log(val_1)
	console.log(val_2)
	
	if (typeof(val_1)=="undefined" ||typeof(val_2)=="undefined")
	{
		alert("Please select Guildline and Grade");
		$('#five_point_zero', row).val("");
		$('#zeor_point_five', row).val("");
		$('#remark', row).val("");
		
		return ;
	}
	
    var zeor_point_five  = $('#zeor_point_five', row).val();
	var five_point_zero   = $('#five_point_zero', row).val();	
	
	if (zeor_point_five<=parseInt(val_1) && five_point_zero <= parseInt(val_2))
	{
			$('#remark', row).val("Pass");
	}	
	else {$('#remark', row).val("Fail");}
	
	
});


function updateCompanyDetails()
{
	code_tbl = document.getElementsByClassName("code_tbl")[0]
	code_rows = code_tbl.rows
	
	for(var j = 1; j<code_rows.length; j++)
	{
		tds = code_rows[j].children
		
		if (tds[0].firstElementChild.value=="")
		{	
				alert("Filter COMPANY NAME cannot be blank in row : "+j);
				return;
		}
		
		if (tds[1].firstElementChild.value=="")
		{	
				alert("Filter ADDRESS cannot be blank in row : "+j);
				return;
		}
		
		
		if (tds[2].firstElementChild.value=="")
		{	
				alert("Filter REPORT NUMBER  cannot be blank in row : "+j);
				return;
		}		
	}
	var final_table_data = {};
    var full_data = {};	
	for(var j = 1; j<code_rows.length; j++)
	{
		tds = code_rows[j].children	
		var table_data = {};
		table_data['COMPANY_NAME'] =tds[0].firstElementChild.value 	
		table_data['ADDRESS'] =tds[1].firstElementChild.value 	
		table_data['REPORT_NUMBER'] =tds[2].firstElementChild.value 		
		final_table_data[j] = table_data
	}
	
	full_data['observation']=final_table_data
	
	$.getJSON('/submit_updateCompanyDetails', 
	{
		params_data : JSON.stringify(full_data)
	}, function(result) 
	{
		alert("done");
		

		
	});
	
}




