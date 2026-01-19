jQuery(document).ready(function($) {
    // Generate Google Apps Script
    $("#ugsiw-generate-script").on("click", function(e) {
        e.preventDefault();
        
        var button = $(this);
        var originalText = button.text();
        
        button.text("Generating...").prop("disabled", true);
        
        // Get selected fields
        var selectedFields = [];
        $("input[name='ugsiw_gs_selected_fields[]']:checked").each(function() {
            selectedFields.push($(this).val());
        });
        
        $.ajax({
            url: ugsiw_gs.ajax_url,
            type: "POST",
            data: {
                action: "ugsiw_generate_google_script",
                nonce: ugsiw_gs.nonce,
                fields: selectedFields
            },
            success: function(response) {
                if (response.success) {
                    $("#ugsiw-generated-script").val(response.data.script);
                    $("#ugsiw-script-output").show();
                    $("html, body").animate({
                        scrollTop: $("#ugsiw-script-output").offset().top - 100
                    }, 500);
                } else {
                    alert("Error generating script. Please try again.");
                }
            },
            error: function() {
                alert("Error generating script. Please try again.");
            },
            complete: function() {
                button.text(originalText).prop("disabled", false);
            }
        });
    });

    // Copy to clipboard
    $("#ugsiw-copy-script").on("click", function() {
        var textarea = $("#ugsiw-generated-script")[0];
        textarea.select();
        document.execCommand("copy");
        
        $("#ugsiw-copy-status").show().fadeOut(2000);
    });

    // Handle required fields
    $("input[name='ugsiw_gs_selected_fields[]']").each(function() {
        if ($(this).is(":disabled")) {
            $(this).prop("checked", true);
        }
    });

    // Prevent unchecking required fields
    $("input[name='ugsiw_gs_selected_fields[]']").on("change", function() {
        if ($(this).is(":disabled") && !$(this).is(":checked")) {
            $(this).prop("checked", true);
        }
    });

    // Toggle fields visibility
    $(".ugsiw-field-category h3").on("click", function() {
        $(this).next(".ugsiw-fields-grid").slideToggle(300);
        $(this).find(".dashicons").toggleClass("dashicons-arrow-down dashicons-arrow-up");
    });

    // Search fields
    $("#ugsiw-field-search").on("keyup", function() {
        var search = $(this).val().toLowerCase();
        $(".ugsiw-field-item").each(function() {
            var label = $(this).find("label").text().toLowerCase();
            if (label.indexOf(search) > -1) {
                $(this).show();
            } else {
                $(this).hide();
            }
        });
    });
});