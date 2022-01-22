var SpecObject = null;
var TestingSpecType = null;
var SpecTestingTimer = {};
var SpecTestingTimerId = 0;
var TestingModuleContent = {};


window.onload = function() {
    SpecObject = new SpecTestingTime();
    SpecObject.init();
    SpecObject.getSpecDeviceModels();
};


window.onclick = function() {

    $("button.test_plan").click(function(event) {
        /*
         * Click 'Spec Type' buttom
         */
         SpecObject.specOnClick(event);
    });

    $("input.test_item").change(function() {
        /*
         * Select every test items.
         */
        SpecObject.calculateTestingTime(this);
    });

    $("#add_testing_item").unbind().click(function() {
        /*
         * Click 'OK' buttom
         * 
         * Note:
         *    One click trigger click-event multiple times.
         *    Answer at StackOverFlow:
         *    https://stackoverflow.com/questions/14969960/jquery-click-events-firing-multiple-times
         */
        console.log("Selected and saved the setting with specific device mmodule.");
        SpecObject.addSelectedTestingItem();
    });

    $(".remove_test_items").unbind().click(function(event) {
        /*
         * Click the buttom at 'x' in tag
         */
        var target_remove_test_item = $(event.target).parent();
        SpecObject.removeTargetTestItem(target_remove_test_item);
    });

    $(".remove_all_test_items").unbind().click(function(event) {
        /*
         * Click 'Clear' buttom
         */
        var target_remove_test_item = $(event.target).parent();
        SpecObject.removeAllTestItem(target_remove_test_item);
    });
    
    $("#export_report").unbind().click(function(event) {
        /*
         * Click 'Export Report' buttom
         */
        console.log("Export report with target spec testing items.");
        // Do something hanlding.
        var target = JSON.stringify(SpecTestingTimer);
        SpecObject.exportSpecTestingTimeReport(target);
    });
    
};


function SpecTestingTime() {
    this.init();    
};


SpecTestingTime.prototype.init = function() {
    // Set headers of table which display target testing items with specific device module
    $("thead#target_test_items_form_theader").append("<th scope='col'></th><th scope='col'> Index </th><th scope='col'> Device Module </th><th scope='col'> Testing Spec Type </th>");
}


SpecTestingTime.prototype.specOnClick = function(event) {
    var self = this;
    TestingSpecType = $(event.target).val()
    // var device_model = $("select#device_model :selected").val();
    // console.error("device_model: " + device_model)
    var device_model = $("select#device_model").val();
    var testing_spec = event.target.textContent;

    $("span#current_device_model_value").text(device_model);
    $("span#current_test_plan_value").text(testing_spec);

    console.error("device_model: " + device_model)
    self.getSpecItemsTestingTime(TestingSpecType, device_model);
}


SpecTestingTime.prototype.calculateTestingTime = function(this_ele) {
    var current_total_time = $("#total_testing_time_value").text();
    var total_testing_time = 0;

    if ($("input:checkbox[class=test_item]:checked").length == 0) {}
    
    $("input:checkbox[class=test_item]:checked").each(function() {
        var testing_time = Object.values(JSON.parse(this.value))[0];
        total_testing_time = parseFloat(total_testing_time) + parseFloat(testing_time);
    });

    $("#total_testing_time_value").empty();
    $("#total_testing_time_value").append(parseFloat(current_total_time) + parseFloat(total_testing_time));
}


SpecTestingTime.prototype.getSpecDeviceModels = function () {
    $.ajax({
        url: "/api/v1/deviceModels", 
        type: "GET", 
        success: function(response) {
            $.each(response, function(index, model) {
                $("#device_model").append("<option value='" + model + "'>" + model + "</option>");
            });
        },
        error: function(xhr) {
            console.log("Occur somemthing unexpected error ...");
        }
    });
}


SpecTestingTime.prototype.getSpecItemsTestingTime = function(spec_type, device_model) {
    var self = this;
    $.ajax({
        url: "/api/v1/testing-time", 
        type: "POST", 
        data: {
            spec: spec_type, 
            deviceModel: device_model
        }, 
        success: function(response) {
            self.displayTestItemsTestingTime(response);
        },
        error: function(xhr) {
            console.log("Occur somemthing unexpected error ...");
        }
    });
}


SpecTestingTime.prototype.displayTestItemsTestingTime = function(response) {
    var self = this;
    // Clean all things in the area
    $("thead#test_items_form_theader").empty();
    $("tbody#test_items_form_tbody").empty();
    
    // Set table headers
    $("thead#test_items_form_theader").append("<th scope='col'> Selected </th><th scope='col'> Test Item </th><th scope='col'> Testing Time </th>");

    // Set table data
    $.each(response, function(index, value) {
        $("tbody#test_items_form_tbody").append("<tr><th scope='row'><input type='checkbox' class='test_item' value='" + self.asJsonType(index, value) + "'></th><td>" + index + "</td><td>" + value + "</td></tr>");
    });
}


SpecTestingTime.prototype.addSelectedTestingItem = function() {debugger
    // Read the setting of specific device, spec and testing items.
    var spec_type = TestingSpecType;
    var device_model = $("select#device_model :selected").val();
    var testing_items = new Array;
    $("input:checkbox[class=test_item]:checked").each(function() {
        testing_items.push(JSON.parse(JSON.stringify($(this).val())));
    });

    // Convert to JSON data
    // Method 1
    var testing_item_info = {
        device_model: device_model,
        spec_type: spec_type,
        items: JSON.parse(JSON.stringify(testing_items))
    };
    // Method 2
    // var testing_item_info = {}
    // testing_item_info.device_model = device_model;
    // testing_item_info.spec_type = spec_type;
    // testing_item_info.items = JSON.parse(JSON.stringify(testing_items));
    
    // Persistent the setting to the variable
    // Note: Should use closure to saving data to be more clear and saver.
    var SpecSettingItemId = "spec_testing_time_" + SpecTestingTimerId
    // Check the target device module and spec type weather is duplicated or not.
    if (SpecTestingTimerId > 0) {
        for (let index = 0; index < SpecTestingTimerId; index ++) {
            var allJsonData = JSON.parse(JSON.stringify(SpecTestingTimer));
            var jdata = JSON.parse(JSON.stringify(allJsonData["spec_testing_time_" + index]));
            if ((jdata.device_model == device_model) && (jdata.spec_type == spec_type)) {
                alert("The device module and Spec type you selected is duplicated. Please edit it, doesn't create it again.");
                return ;
            }
        }
    }
    // Save into object
    SpecTestingTimer[SpecSettingItemId] = testing_item_info;

    all_target_modules = Object.keys(TestingModuleContent);
    if (all_target_modules.includes(device_model) == false) {
        // Handling content data
        var content = {};
        content[spec_type] = testing_items;
        TestingModuleContent[device_model] = content;

        // Add <div> element for one specific device module 
        this.generateDeviceModuleArea(device_model)
        // Add the Module name 
        this.generateDeviceModuleTitle(device_model);
        // Add spec title
        this.generateSpecTitleOfDeviceModuleContent(device_model, spec_type);
        // Add detail content about testing item and testing time
        this.generateDeviceModuleContent(device_model, spec_type, testing_items);
    } else {
        if (Object.keys(TestingModuleContent[device_model]).includes(spec_type) == false) {
            // Handling content data
            var content = {};
            content[spec_type] = testing_items;
            TestingModuleContent[device_model] = content;

            this.generateSpecTitleOfDeviceModuleContent(device_model, spec_type);
            this.generateDeviceModuleContent(device_model, spec_type, testing_items);
        } else {
            TestingModuleContent[device_model] = content;
            var total_content = content.concat(testing_items);
            var total_content_set = Array.from(new Set(total_content));
            TestingModuleContent[device_model] = total_content_set;

            this.generateDeviceModuleContent(device_model, spec_type, testing_items);
        }
    }

    // Add the ID value
    SpecTestingTimerId += 1;
}


SpecTestingTime.prototype.formatChar = function(original_char) {
    original_char_array = original_char.split("_");
    char_array = new Array();
    original_char_array.forEach((char_ele) => {
        var first_char = char_ele[0].toUpperCase();
        var left_char = char_ele.substring(1, char_ele.length);
        char_array.push(first_char + left_char);
    })
    return char_array.join(" ");
}


SpecTestingTime.prototype.generateDeviceModuleArea = function(device_model) {
    var area = ' \
        <div id="' + device_model + '_name"></div> \
        <div id="' + device_model + '_content"></div> \
    ';

    $("#device_modules_target_items").append(area);
}


SpecTestingTime.prototype.generateDeviceModuleTitle = function(device_model) {
    var device_model_str = this.formatChar(device_model);
    var deviceModuleTitle = ' \
        <p> \
            <button class="btn" type="button" data-bs-toggle="collapse" data-bs-target="#' + device_model + '_testing_content" aria-expanded="false" aria-controls="' + device_model + '_testing_content"> \
                <span class="glyphicon glyphicon-list"></span> \
                <span class="device_module_name fs-2 fw-bold fst-italic">' + device_model_str + '</span> \
            </button> \
            <span class="remove_all_test_items badge rounded-pill bg-danger"> \
                <span id="tag_content">Clear</span> \
                <span class="glyphicon glyphicon-remove"></span> \
            </span> \
        </p> \
        ';

    $("#" + device_model + "_name").append(deviceModuleTitle);
    // return deviceModuleTitle;
}


SpecTestingTime.prototype.generateSpecTitleOfDeviceModuleContent = function(device_model, spec_type) {
    var spec_type_str = this.formatChar(spec_type);

    var content = ' \
    <div class="collapse" id="' + device_model + '_testing_content"> \
        <div class="card-body" id="' + device_model + '_card_body"> \
            <p class="h3">' + spec_type_str + '</p> \
            <hr size="1" style="border: 1px black solid; margin-top: 0%;"> \
            <div class="testing_items_with_spec" id="' + device_model + "_" + spec_type + '"> \
            </div> \
        </div> \
    </div> \
    ';

    $("#" + device_model + "_content").append(content);
    // return content;
}


SpecTestingTime.prototype.generateDeviceModuleContent = function(device_model, spec_type, items) {
    var content = ' \
    <span class="badge rounded-pill bg-primary"> \
        <span id="tag_content">TEST_ITEM_NAME: TEST_ITEM_TIME</span> \
        <span class="remove_test_items glyphicon glyphicon-remove"></span> \
    </span> \
    ';

    items.forEach((test_item) => {
        var test_item_json = JSON.parse(test_item);
        var test_item_name = Object.keys(test_item_json)[0];
        var test_item_time = test_item_json[test_item_name];
        var content_has_item = content.replace("TEST_ITEM_NAME", test_item_name);
        var content_has_item_and_time = content_has_item.replace("TEST_ITEM_TIME", test_item_time);
        $("#" + device_model + "_" + spec_type).append(content_has_item_and_time);
    })

    // return content;
}


SpecTestingTime.prototype.removeTargetTestItem = function(element) {
    // Get the target test item
    var test_item = $(element).text().replaceAll(" ", "").split(":")[0];
    var module_and_spec = $(element).parent().attr("id");
    var spec_re = /_.{1,16}_test/g;
    var underline_spec_type = module_and_spec.match(spec_re)[0];
    // Get the Spec type of target test item
    var spec_type = underline_spec_type.substring(1);
    // Get the device module of target test item
    var device_module = module_and_spec.split(underline_spec_type)[0];

    var test_items = TestingModuleContent[device_module][spec_type];
    var new_test_items = test_items.filter(function(value, index, arr) {
        return value.match(test_item) == null;
    });
    TestingModuleContent[device_module][spec_type] = new_test_items;

    $(element).remove();
}


SpecTestingTime.prototype.removeAllTestItem = function(element) {
    // Delete all test items
    var target_remove_all_device_module = $(element).prev().text().replaceAll(" ", "")
    delete TestingModuleContent[target_remove_all_device_module];

    // Remove all element location
    var module_title = target_remove_all_device_module + "_name";
    var module_content = target_remove_all_device_module + "_content";
    $("#" + module_title).remove();
    $("#" + module_content).remove();
}


SpecTestingTime.prototype.asJsonType = function(index, value) {
    var testing_item = index.replaceAll(' ', '_');
    return '{"' + testing_item + '": ' + value + '}';
}


SpecTestingTime.prototype.saveSpecTestingTimeCalculationSetting = function() {debugger
    // Read the setting of specific device, spec and testing items.
    var spec_type = TestingSpecType;
    var device_model = $("select#device_model :selected").val();
    var testing_items = new Array;
    $("input:checkbox[class=test_item]:checked").each(function() {
        testing_items.push(JSON.parse(JSON.stringify($(this).val())));
    });

    // Convert to JSON data
    // Method 1
    var testing_item_info = {
        device_model: device_model,
        spec_type: spec_type,
        items: JSON.parse(JSON.stringify(testing_items))
    };
    // Method 2
    // var testing_item_info = {}
    // testing_item_info.device_model = device_model;
    // testing_item_info.spec_type = spec_type;
    // testing_item_info.items = JSON.parse(JSON.stringify(testing_items));
    
    // Persistent the setting to the variable
    // Note: Should use closure to saving data to be more clear and saver.
    var SpecSettingItemId = "spec_testing_time_" + SpecTestingTimerId
    // Check the target device module and spec type weather is duplicated or not.
    if (SpecTestingTimerId > 0) {
        for (let index = 0; index < SpecTestingTimerId; index ++) {
            var allJsonData = JSON.parse(JSON.stringify(SpecTestingTimer));
            var jdata = JSON.parse(JSON.stringify(allJsonData["spec_testing_time_" + index]));
            if ((jdata.device_model == device_model) && (jdata.spec_type == spec_type)) {
                alert("The device module and Spec type you selected is duplicated. Please edit it, doesn't create it again.");
                return ;
            }
        }
    }
    // Save into object
    SpecTestingTimer[SpecSettingItemId] = testing_item_info;

    // Set table data
    $("tbody#target_test_items_form_tbody").append("<tr><th scope='row'><input type='checkbox' class='test_item' value='" + SpecSettingItemId + "'><th scope='row'><span>" + SpecTestingTimerId + "</span></th><td>" + device_model + "</td><td>" + spec_type + "</td></tr>");

    // Add the ID value
    SpecTestingTimerId += 1;
}


SpecTestingTime.prototype.backToEditor = function() {
    /*
     * Note: (Not implement)
     *     Back to edit step to review or edit the setting with specific device module & spec type
     */
}


SpecTestingTime.prototype.exportSpecTestingTimeReport = function(target_setting) {
    /*
     * Note:
     * Refer to URL https://www.aspsnippets.com/Articles/Download-File-in-AJAX-Response-Success-using-jQuery.aspx
     */
    $.ajax({
        url: "/api/v1/export-testing-time-result", 
        type: "POST", 
        data: {
            target: target_setting
        },
        cache: false,
        xhr: function () {
            var xhr = new XMLHttpRequest();
            xhr.onreadystatechange = function () {
                if (xhr.readyState == 2) {
                    if (xhr.status == 200) {
                        xhr.responseType = "blob";
                    } else {
                        xhr.responseType = "text";
                    }
                }
            };
            return xhr;
        },
        success: function(response) {
            const blob = new Blob([response], {type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;"});
            const downloadUrl = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = downloadUrl;
            a.download = "D-Link_Wi-Fi_Testing_Time_Report_.xlsx";
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
        },
        error: function(xhr) {
            console.log(xhr);
            console.log("Occur somemthing unexpected error ...");
        }
    });
}

