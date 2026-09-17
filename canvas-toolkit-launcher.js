javascript:(function(){
function runSyllabusBuilder(){
  var m = window.location.pathname.match(/\/courses\/(\d+)/);
  if(!m){ alert('Run this bookmarklet from inside a Canvas course (a URL like /courses/12345).'); return; }
  var courseId = m[1];
  var base = window.location.origin;

  function esc(s){
    return (s || '').replace(/[&<>]/g, function(c){ return {'&':'&amp;','<':'&lt;','>':'&gt;'}[c]; });
  }
  function formatDate(d){
    var dt = new Date(d);
    return dt.toLocaleDateString('en-US', {weekday:'short', month:'short', day:'numeric', year:'numeric'});
  }
  function mondayOf(d){
    var dt = new Date(d);
    var day = dt.getDay();
    var diff = (day === 0 ? -6 : 1 - day);
    var monday = new Date(dt);
    monday.setDate(dt.getDate() + diff);
    monday.setHours(0,0,0,0);
    return monday;
  }
  function weekLabel(monday){
    var sunday = new Date(monday);
    sunday.setDate(monday.getDate() + 6);
    var opts = {month:'short', day:'numeric'};
    var startStr = monday.toLocaleDateString('en-US', opts);
    var endStr = sunday.toLocaleDateString('en-US', {month:'short', day:'numeric', year:'numeric'});
    return 'Week of ' + startStr + ' &ndash; ' + endStr;
  }
  function fetchJSON(url){
    return fetch(url, {credentials:'same-origin', headers:{'Accept':'application/json'}}).then(function(r){
      if(!r.ok) throw new Error('Canvas API request failed (' + r.status + '): ' + url);
      return r.json();
    });
  }
  function showOverlay(html){
    var old = document.getElementById('sb-overlay');
    if(old) old.remove();
    var overlay = document.createElement('div');
    overlay.id = 'sb-overlay';
    overlay.style.cssText = 'position:fixed;top:16px;left:16px;right:16px;bottom:16px;background:#fff;z-index:999999;overflow:auto;padding:24px;border:1px solid #ccc;border-radius:6px;box-shadow:0 4px 24px rgba(0,0,0,0.35);font-family:-apple-system,Helvetica,Arial,sans-serif;color:#222;';
    overlay.innerHTML =
      '<div style="text-align:right;margin-bottom:14px;position:sticky;top:0;background:#fff;padding-bottom:8px;border-bottom:1px solid #eee;">' +
        '<button id="sb-copy-btn" style="padding:8px 14px;margin-right:8px;cursor:pointer;border:1px solid #999;border-radius:4px;background:#f5f5f5;">Copy HTML</button>' +
        '<button id="sb-word-btn" style="padding:8px 14px;margin-right:8px;cursor:pointer;border:1px solid #999;border-radius:4px;background:#f5f5f5;">Download Word Doc</button>' +
        '<button id="sb-close-btn" style="padding:8px 14px;cursor:pointer;border:1px solid #999;border-radius:4px;background:#f5f5f5;">Close</button>' +
      '</div>' +
      '<div id="sb-content">' + html + '</div>';
    document.body.appendChild(overlay);
    document.getElementById('sb-close-btn').onclick = function(){ overlay.remove(); };
    document.getElementById('sb-copy-btn').onclick = function(){
      var range = document.createRange();
      range.selectNode(document.getElementById('sb-content'));
      var sel = window.getSelection();
      sel.removeAllRanges();
      sel.addRange(range);
      document.execCommand('copy');
      sel.removeAllRanges();
      alert('Copied. In the Syllabus page editor, switch the Rich Content Editor to HTML view and paste there for the cleanest result.');
    };
    document.getElementById('sb-word-btn').onclick = function(){
      var content = document.getElementById('sb-content').innerHTML;
      var docHtml = '<html xmlns:o="urn:schemas-microsoft-com:office:office" ' +
        'xmlns:w="urn:schemas-microsoft-com:office:word" xmlns="http://www.w3.org/TR/REC-html40">' +
        '<head><meta charset="utf-8"><title>Syllabus Export</title></head><body>' + content + '</body></html>';
      var blob = new Blob(['\ufeff', docHtml], {type: 'application/msword'});
      var url = URL.createObjectURL(blob);
      var a = document.createElement('a');
      a.href = url;
      a.download = 'course-' + courseId + '-syllabus.doc';
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      setTimeout(function(){ URL.revokeObjectURL(url); }, 1000);
    };
  }

  Promise.all([
    fetchJSON(base + '/api/v1/courses/' + courseId + '?include[]=syllabus_body'),
    fetchJSON(base + '/api/v1/courses/' + courseId + '/modules?include[]=items&per_page=100'),
    fetchJSON(base + '/api/v1/courses/' + courseId + '/assignments?per_page=100'),
    fetchJSON(base + '/api/v1/courses/' + courseId + '/assignment_groups?per_page=100')
  ]).then(function(results){
    var course = results[0], modules = results[1], assignments = results[2], groups = results[3];
    var weighted = course.apply_assignment_group_weights === true;

    var publishedAssignments = assignments.filter(function(a){ return a.published !== false; });

    var byAssignmentId = {}, byQuizId = {}, byDiscussionTopicId = {};
    publishedAssignments.forEach(function(a){
      byAssignmentId[a.id] = a;
      if(a.quiz_id) byQuizId[a.quiz_id] = a;
      if(a.discussion_topic && a.discussion_topic.id) byDiscussionTopicId[a.discussion_topic.id] = a;
    });

    var gradableWithDue = [];
    var gradableNoDue = [];
    var resourcesByModule = {};
    var unpublishedModules = [];
    var unpublishedItems = [];

    modules.forEach(function(mod){
      if(mod.published === false){
        unpublishedModules.push(mod.name);
        return;
      }
      (mod.items || []).forEach(function(item){
        if(item.published === false){
          unpublishedItems.push({module: mod.name, item: item.title});
          return;
        }
        var matched = null;
        if(item.type === 'Assignment') matched = byAssignmentId[item.content_id];
        else if(item.type === 'Quiz') matched = byQuizId[item.content_id];
        else if(item.type === 'Discussion') matched = byDiscussionTopicId[item.content_id];

        if(matched){
          if(matched.due_at){
            gradableWithDue.push({title: item.title, due_at: new Date(matched.due_at), points: matched.points_possible || 0});
          } else {
            gradableNoDue.push({title: item.title, points: matched.points_possible || 0});
          }
        } else if(item.type === 'Assignment' || item.type === 'Quiz' || item.type === 'Discussion'){
          gradableNoDue.push({title: item.title, points: 0});
        } else {
          if(!resourcesByModule[mod.name]) resourcesByModule[mod.name] = [];
          resourcesByModule[mod.name].push(item.title);
        }
      });
    });

    var weeks = {};
    gradableWithDue.forEach(function(g){
      var monday = mondayOf(g.due_at);
      var key = monday.toISOString().slice(0,10);
      if(!weeks[key]) weeks[key] = {monday: monday, items: []};
      weeks[key].items.push(g);
    });
    var weekKeys = Object.keys(weeks).sort();

    var scheduleHtml = '';
    weekKeys.forEach(function(key){
      var wk = weeks[key];
      wk.items.sort(function(a,b){ return a.due_at - b.due_at; });
      scheduleHtml += '<h3>' + weekLabel(wk.monday) + '</h3><ul>';
      wk.items.forEach(function(g){
        scheduleHtml += '<li>' + esc(g.title) + ' &mdash; Due ' + formatDate(g.due_at) + ' (' + g.points + ' pts)</li>';
      });
      scheduleHtml += '</ul>';
    });
    if(gradableNoDue.length){
      scheduleHtml += '<h3>No Due Date</h3><ul>';
      gradableNoDue.forEach(function(g){
        scheduleHtml += '<li>' + esc(g.title) + ' (' + g.points + ' pts)</li>';
      });
      scheduleHtml += '</ul>';
    }

    var resourcesHtml = '';
    Object.keys(resourcesByModule).forEach(function(modName){
      resourcesHtml += '<h3>' + esc(modName) + '</h3><ul>';
      resourcesByModule[modName].forEach(function(title){
        resourcesHtml += '<li>' + esc(title) + '</li>';
      });
      resourcesHtml += '</ul>';
    });
    if(!resourcesHtml) resourcesHtml = '<p><em>None found.</em></p>';

    var sortedAssignments = publishedAssignments.slice().sort(function(a, b){
      return new Date(a.due_at || '2100-01-01') - new Date(b.due_at || '2100-01-01');
    });
    var asgHtml = '<table border="1" cellpadding="6" style="border-collapse:collapse;width:100%;font-size:14px;">' +
      '<tr><th>Assignment</th><th>Points</th><th>Due Date</th></tr>';
    sortedAssignments.forEach(function(a){
      asgHtml += '<tr><td>' + esc(a.name) + '</td><td>' + (a.points_possible || 0) + '</td><td>' +
        (a.due_at ? formatDate(a.due_at) : 'No due date') + '</td></tr>';
    });
    asgHtml += '</table>';

    var pointsByGroup = {};
    publishedAssignments.forEach(function(a){
      var gid = a.assignment_group_id;
      pointsByGroup[gid] = (pointsByGroup[gid] || 0) + (a.points_possible || 0);
    });
    var gradingHtml = '';
    if(weighted){
      gradingHtml = '<table border="1" cellpadding="6" style="border-collapse:collapse;width:100%;font-size:14px;">' +
        '<tr><th>Category</th><th>Weight</th></tr>';
      groups.forEach(function(g){
        gradingHtml += '<tr><td>' + esc(g.name) + '</td><td>' + (g.group_weight != null ? g.group_weight + '%' : '&mdash;') + '</td></tr>';
      });
      gradingHtml += '</table>';
    } else {
      gradingHtml = '<p><em>This course grades by total points, not weighted categories.</em></p>' +
        '<table border="1" cellpadding="6" style="border-collapse:collapse;width:100%;font-size:14px;">' +
        '<tr><th>Category</th><th>Total Points</th></tr>';
      groups.forEach(function(g){
        gradingHtml += '<tr><td>' + esc(g.name) + '</td><td>' + (pointsByGroup[g.id] || 0) + '</td></tr>';
      });
      gradingHtml += '</table>';
    }

    var unpublishedHtml = '';
    if(unpublishedModules.length || unpublishedItems.length){
      if(unpublishedModules.length){
        unpublishedHtml += '<p><strong>Unpublished modules:</strong></p><ul>';
        unpublishedModules.forEach(function(name){ unpublishedHtml += '<li>' + esc(name) + '</li>'; });
        unpublishedHtml += '</ul>';
      }
      if(unpublishedItems.length){
        unpublishedHtml += '<p><strong>Unpublished items in published modules:</strong></p><ul>';
        unpublishedItems.forEach(function(u){ unpublishedHtml += '<li>' + esc(u.module) + ' &mdash; ' + esc(u.item) + '</li>'; });
        unpublishedHtml += '</ul>';
      }
    } else {
      unpublishedHtml = '<p><em>Nothing unpublished found.</em></p>';
    }

    var fullHtml = '<h2>Weekly Schedule</h2>' + (scheduleHtml || '<p><em>No graded items with due dates found.</em></p>') +
      '<h2>Additional Course Content (no due date)</h2>' + resourcesHtml +
      '<h2>Assignment List</h2>' + asgHtml +
      '<h2>Grading Breakdown</h2>' + gradingHtml +
      '<h2>Unpublished (not in syllabus)</h2>' + unpublishedHtml +
      '<p style="color:#888;font-size:12px;">Generated from course ID ' + courseId + '. Pulls at most 100 items per category &mdash; larger courses may need manual review of anything beyond that.</p>';

    showOverlay(fullHtml);
  }).catch(function(err){
    alert('Could not build the syllabus summary: ' + err.message);
  });
}

function runClassRosterBuilder(){
  var base = window.location.origin;

  function esc(s){
    return (s || '').replace(/[&<>]/g, function(c){ return {'&':'&amp;','<':'&lt;','>':'&gt;'}[c]; });
  }
  function xesc(s){
    return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  }
  function formatDateTime(d){
    if(!d) return 'Never';
    var dt = new Date(d);
    return dt.toLocaleDateString('en-US', {month:'short', day:'numeric', year:'numeric'}) +
      ' ' + dt.toLocaleTimeString('en-US', {hour:'numeric', minute:'2-digit'});
  }
  function formatGrade(enrollment){
    if(!enrollment || !enrollment.grades) return 'N/A';
    var g = enrollment.grades;
    if(g.current_score != null) return g.current_score + '%';
    if(g.current_grade) return g.current_grade;
    return 'N/A';
  }
  function fetchJSON(url){
    return fetch(url, {credentials:'same-origin', headers:{'Accept':'application/json'}}).then(function(r){
      if(!r.ok) throw new Error('Canvas API request failed (' + r.status + '): ' + url);
      return r.json();
    });
  }
  function removeOverlay(id){
    var old = document.getElementById(id);
    if(old) old.remove();
  }
  function baseOverlay(id){
    removeOverlay(id);
    var overlay = document.createElement('div');
    overlay.id = id;
    overlay.style.cssText = 'position:fixed;top:16px;left:16px;right:16px;bottom:16px;background:#fff;z-index:999999;overflow:auto;padding:24px;border:1px solid #ccc;border-radius:6px;box-shadow:0 4px 24px rgba(0,0,0,0.35);font-family:-apple-system,Helvetica,Arial,sans-serif;color:#222;';
    document.body.appendChild(overlay);
    return overlay;
  }
  function sanitizeSheetName(name, used){
    var clean = String(name).replace(/[:\\\/\?\*\[\]]/g, ' ').trim();
    if(clean.length > 31) clean = clean.slice(0, 31);
    if(!clean) clean = 'Sheet';
    var candidate = clean, n = 2;
    while(used[candidate]){
      var suffix = ' (' + n + ')';
      candidate = clean.slice(0, 31 - suffix.length) + suffix;
      n++;
    }
    used[candidate] = true;
    return candidate;
  }

  function buildXlsAndOverlay(termName, courseRosters){
    var overlay = baseOverlay('crb-overlay');
    var bodyHtml = '<div style="text-align:right;margin-bottom:14px;position:sticky;top:0;background:#fff;padding-bottom:8px;border-bottom:1px solid #eee;">' +
      '<button id="crb-xls-btn" style="padding:8px 14px;margin-right:8px;cursor:pointer;border:1px solid #999;border-radius:4px;background:#f5f5f5;">Download .xls</button>' +
      '<button id="crb-close-btn" style="padding:8px 14px;cursor:pointer;border:1px solid #999;border-radius:4px;background:#f5f5f5;">Close</button>' +
      '</div><div id="crb-content">';

    bodyHtml += '<h2>' + esc(termName) + ' &mdash; Class Rosters</h2>';

    courseRosters.forEach(function(cr){
      bodyHtml += '<h3>' + esc(cr.courseName) + (cr.courseCode ? ' (' + esc(cr.courseCode) + ')' : '') + '</h3>';
      if(cr.error){
        bodyHtml += '<p><em>Could not load roster: ' + esc(cr.error) + '</em></p>';
        return;
      }
      if(!cr.students.length){
        bodyHtml += '<p><em>No students found.</em></p>';
        return;
      }
      bodyHtml += '<table border="1" cellpadding="6" style="border-collapse:collapse;width:100%;font-size:14px;margin-bottom:20px;">' +
        '<tr><th>Name</th><th>Email</th><th>Section</th><th>Last Access</th><th>Current Grade</th></tr>';
      cr.students.forEach(function(s){
        bodyHtml += '<tr><td>' + esc(s.name) + '</td><td>' + esc(s.email || '') + '</td><td>' + esc(s.section || '') +
          '</td><td>' + esc(s.lastAccess) + '</td><td>' + esc(s.grade) + '</td></tr>';
      });
      bodyHtml += '</table>';
    });
    bodyHtml += '</div>';
    overlay.innerHTML = bodyHtml;

    document.getElementById('crb-close-btn').onclick = function(){ overlay.remove(); };
    document.getElementById('crb-xls-btn').onclick = function(){
      var usedNames = {};
      var xml = '<?xml version="1.0"?>' +
        '<?mso-application progid="Excel.Sheet"?>' +
        '<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet" ' +
        'xmlns:o="urn:schemas-microsoft-com:office:office" ' +
        'xmlns:x="urn:schemas-microsoft-com:office:excel" ' +
        'xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet" ' +
        'xmlns:html="http://www.w3.org/TR/REC-html40">' +
        '<Styles><Style ss:ID="Header"><Font ss:Bold="1"/></Style></Styles>';

      var summaryName = sanitizeSheetName('Summary', usedNames);
      xml += '<Worksheet ss:Name="' + summaryName + '"><Table>' +
        '<Row>' +
          '<Cell ss:StyleID="Header"><Data ss:Type="String">Course</Data></Cell>' +
          '<Cell ss:StyleID="Header"><Data ss:Type="String">Course Code</Data></Cell>' +
          '<Cell ss:StyleID="Header"><Data ss:Type="String">Student Count</Data></Cell>' +
        '</Row>';
      courseRosters.forEach(function(cr){
        xml += '<Row>' +
          '<Cell><Data ss:Type="String">' + xesc(cr.courseName) + '</Data></Cell>' +
          '<Cell><Data ss:Type="String">' + xesc(cr.courseCode || '') + '</Data></Cell>' +
          '<Cell><Data ss:Type="Number">' + (cr.error ? 0 : cr.students.length) + '</Data></Cell>' +
        '</Row>';
      });
      xml += '</Table></Worksheet>';

      courseRosters.forEach(function(cr){
        var sheetName = sanitizeSheetName(cr.courseCode || cr.courseName, usedNames);
        xml += '<Worksheet ss:Name="' + sheetName + '"><Table>' +
          '<Row>' +
            '<Cell ss:StyleID="Header"><Data ss:Type="String">Name</Data></Cell>' +
            '<Cell ss:StyleID="Header"><Data ss:Type="String">Email</Data></Cell>' +
            '<Cell ss:StyleID="Header"><Data ss:Type="String">Section</Data></Cell>' +
            '<Cell ss:StyleID="Header"><Data ss:Type="String">Last Access</Data></Cell>' +
            '<Cell ss:StyleID="Header"><Data ss:Type="String">Current Grade</Data></Cell>' +
          '</Row>';
        (cr.students || []).forEach(function(s){
          xml += '<Row>' +
            '<Cell><Data ss:Type="String">' + xesc(s.name) + '</Data></Cell>' +
            '<Cell><Data ss:Type="String">' + xesc(s.email || '') + '</Data></Cell>' +
            '<Cell><Data ss:Type="String">' + xesc(s.section || '') + '</Data></Cell>' +
            '<Cell><Data ss:Type="String">' + xesc(s.lastAccess) + '</Data></Cell>' +
            '<Cell><Data ss:Type="String">' + xesc(s.grade) + '</Data></Cell>' +
          '</Row>';
        });
        xml += '</Table></Worksheet>';
      });
      xml += '</Workbook>';

      var blob = new Blob([xml], {type: 'application/vnd.ms-excel'});
      var url = URL.createObjectURL(blob);
      var a = document.createElement('a');
      a.href = url;
      a.download = 'class-list-' + termName.replace(/[^a-z0-9]+/gi, '-').toLowerCase() + '.xls';
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      setTimeout(function(){ URL.revokeObjectURL(url); }, 1000);
    };
  }

  function loadRosters(termName, matchingCourses){
    var overlay = baseOverlay('crb-overlay');
    overlay.innerHTML = '<p>Loading rosters for ' + matchingCourses.length + ' course(s) in ' + esc(termName) + '&hellip;</p>';

    Promise.all(matchingCourses.map(function(course){
      return Promise.all([
        fetchJSON(base + '/api/v1/courses/' + course.id + '/users?enrollment_type[]=student&include[]=email&include[]=enrollments&per_page=100'),
        fetchJSON(base + '/api/v1/courses/' + course.id + '/sections?per_page=100')
      ]).then(function(res){
        var users = res[0], sections = res[1];
        var sectionMap = {};
        sections.forEach(function(sec){ sectionMap[sec.id] = sec.name; });
        var students = users.map(function(u){
          var enrollment = (u.enrollments || []).filter(function(e){ return e.type === 'StudentEnrollment'; })[0] || (u.enrollments || [])[0];
          var sectionName = '';
          if(enrollment && sectionMap[enrollment.course_section_id]){
            sectionName = sectionMap[enrollment.course_section_id];
          }
          return {
            name: u.sortable_name || u.name,
            email: u.email,
            section: sectionName,
            lastAccess: formatDateTime(enrollment && enrollment.last_activity_at),
            grade: formatGrade(enrollment)
          };
        }).sort(function(a, b){ return (a.name || '').localeCompare(b.name || ''); });
        return { courseName: course.name, courseCode: course.course_code, students: students };
      }).catch(function(err){
        return { courseName: course.name, courseCode: course.course_code, students: [], error: err.message };
      });
    })).then(function(courseRosters){
      buildXlsAndOverlay(termName, courseRosters);
    });
  }

  function showTermPicker(courses){
    var termNames = {};
    courses.forEach(function(c){
      if(c.term && c.term.name) termNames[c.term.name] = true;
    });
    var names = Object.keys(termNames).sort();
    if(!names.length){
      alert('No terms found on your courses.');
      return;
    }
    var overlay = baseOverlay('crb-overlay');
    var html = '<h2>Pick a term</h2><div>';
    names.forEach(function(n, i){
      html += '<button class="crb-term-btn" data-idx="' + i + '" style="display:block;width:100%;text-align:left;padding:10px 14px;margin-bottom:8px;cursor:pointer;border:1px solid #999;border-radius:4px;background:#f5f5f5;">' + esc(n) + '</button>';
    });
    html += '</div><button id="crb-close-btn" style="margin-top:10px;padding:8px 14px;cursor:pointer;border:1px solid #999;border-radius:4px;background:#f5f5f5;">Cancel</button>';
    overlay.innerHTML = html;
    document.getElementById('crb-close-btn').onclick = function(){ overlay.remove(); };
    var buttons = overlay.querySelectorAll('.crb-term-btn');
    buttons.forEach(function(btn){
      btn.onclick = function(){
        var termName = names[parseInt(btn.getAttribute('data-idx'), 10)];
        var matching = courses.filter(function(c){ return c.term && c.term.name === termName; });
        loadRosters(termName, matching);
      };
    });
  }

  fetchJSON(base + '/api/v1/users/self/courses?enrollment_type=teacher&include[]=term&include[]=total_students&per_page=100&state[]=unpublished&state[]=available&state[]=completed')
    .then(function(courses){
      showTermPicker(courses);
    }).catch(function(err){
      alert('Could not load your courses: ' + err.message);
    });
}

function runAtRiskDashboard(){
  var base = window.location.origin;
  var INACTIVE_DAYS = 7;
  var GRADE_THRESHOLD = 70;
  var MISSING_THRESHOLD = 2;

  function esc(s){
    return (s || '').replace(/[&<>]/g, function(c){ return {'&':'&amp;','<':'&lt;','>':'&gt;'}[c]; });
  }
  function xesc(s){
    return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  }
  function fetchJSON(url){
    return fetch(url, {credentials:'same-origin', headers:{'Accept':'application/json'}}).then(function(r){
      if(!r.ok) throw new Error('Canvas API request failed (' + r.status + '): ' + url);
      return r.json();
    });
  }
  function removeOverlay(id){
    var old = document.getElementById(id);
    if(old) old.remove();
  }
  function baseOverlay(id){
    removeOverlay(id);
    var overlay = document.createElement('div');
    overlay.id = id;
    overlay.style.cssText = 'position:fixed;top:16px;left:16px;right:16px;bottom:16px;background:#fff;z-index:999999;overflow:auto;padding:24px;border:1px solid #ccc;border-radius:6px;box-shadow:0 4px 24px rgba(0,0,0,0.35);font-family:-apple-system,Helvetica,Arial,sans-serif;color:#222;';
    document.body.appendChild(overlay);
    return overlay;
  }
  function sanitizeSheetName(name, used){
    var clean = String(name).replace(/[:\\\/\?\*\[\]]/g, ' ').trim();
    if(clean.length > 31) clean = clean.slice(0, 31);
    if(!clean) clean = 'Sheet';
    var candidate = clean, n = 2;
    while(used[candidate]){
      var suffix = ' (' + n + ')';
      candidate = clean.slice(0, 31 - suffix.length) + suffix;
      n++;
    }
    used[candidate] = true;
    return candidate;
  }
  function daysSince(dateStr){
    if(!dateStr) return null;
    var diffMs = Date.now() - new Date(dateStr).getTime();
    return Math.floor(diffMs / (1000 * 60 * 60 * 24));
  }
  function formatDateTime(d){
    if(!d) return 'Never';
    var dt = new Date(d);
    return dt.toLocaleDateString('en-US', {month:'short', day:'numeric', year:'numeric'}) +
      ' ' + dt.toLocaleTimeString('en-US', {hour:'numeric', minute:'2-digit'});
  }

  function buildXlsAndOverlay(termName, courseReports){
    var overlay = baseOverlay('arb-overlay');
    var bodyHtml = '<div style="text-align:right;margin-bottom:14px;position:sticky;top:0;background:#fff;padding-bottom:8px;border-bottom:1px solid #eee;">' +
      '<button id="arb-xls-btn" style="padding:8px 14px;margin-right:8px;cursor:pointer;border:1px solid #999;border-radius:4px;background:#f5f5f5;">Download .xls</button>' +
      '<button id="arb-close-btn" style="padding:8px 14px;cursor:pointer;border:1px solid #999;border-radius:4px;background:#f5f5f5;">Close</button>' +
      '</div><div id="arb-content">';

    bodyHtml += '<h2>' + esc(termName) + ' &mdash; At-Risk Early Warning</h2>' +
      '<p style="color:#666;font-size:13px;">Flags: inactive ' + INACTIVE_DAYS + '+ days, grade below ' + GRADE_THRESHOLD + '%, or ' + MISSING_THRESHOLD + '+ missing assignments.</p>';

    courseReports.forEach(function(cr){
      bodyHtml += '<h3>' + esc(cr.courseName) + (cr.courseCode ? ' (' + esc(cr.courseCode) + ')' : '') + '</h3>';
      if(cr.error){
        bodyHtml += '<p><em>Could not load data: ' + esc(cr.error) + '</em></p>';
        return;
      }
      if(cr.missingUnavailable){
        bodyHtml += '<p style="color:#a66;font-size:13px;"><em>Analytics not available for this course &mdash; missing-assignment flag skipped.</em></p>';
      }
      bodyHtml += '<p>' + cr.atRisk.length + ' of ' + cr.totalStudents + ' students flagged.</p>';
      if(!cr.atRisk.length){
        bodyHtml += '<p><em>No students currently flagged.</em></p>';
        return;
      }
      bodyHtml += '<table border="1" cellpadding="6" style="border-collapse:collapse;width:100%;font-size:14px;margin-bottom:20px;">' +
        '<tr><th>Name</th><th>Email</th><th>Section</th><th>Last Access</th><th>Current Grade</th><th>Missing</th><th>Flags</th></tr>';
      cr.atRisk.forEach(function(s){
        bodyHtml += '<tr><td>' + esc(s.name) + '</td><td>' + esc(s.email || '') + '</td><td>' + esc(s.section || '') +
          '</td><td>' + esc(s.lastAccessDisplay) + '</td><td>' + esc(s.gradeDisplay) + '</td><td>' + esc(s.missingDisplay) +
          '</td><td>' + esc(s.reasons.join('; ')) + '</td></tr>';
      });
      bodyHtml += '</table>';
    });
    bodyHtml += '</div>';
    overlay.innerHTML = bodyHtml;

    document.getElementById('arb-close-btn').onclick = function(){ overlay.remove(); };
    document.getElementById('arb-xls-btn').onclick = function(){
      var usedNames = {};
      var xml = '<?xml version="1.0"?>' +
        '<?mso-application progid="Excel.Sheet"?>' +
        '<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet" ' +
        'xmlns:o="urn:schemas-microsoft-com:office:office" ' +
        'xmlns:x="urn:schemas-microsoft-com:office:excel" ' +
        'xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet" ' +
        'xmlns:html="http://www.w3.org/TR/REC-html40">' +
        '<Styles><Style ss:ID="Header"><Font ss:Bold="1"/></Style></Styles>';

      var summaryName = sanitizeSheetName('Summary', usedNames);
      xml += '<Worksheet ss:Name="' + summaryName + '"><Table>' +
        '<Row>' +
          '<Cell ss:StyleID="Header"><Data ss:Type="String">Course</Data></Cell>' +
          '<Cell ss:StyleID="Header"><Data ss:Type="String">Course Code</Data></Cell>' +
          '<Cell ss:StyleID="Header"><Data ss:Type="String">Total Students</Data></Cell>' +
          '<Cell ss:StyleID="Header"><Data ss:Type="String">At-Risk Count</Data></Cell>' +
        '</Row>';
      courseReports.forEach(function(cr){
        xml += '<Row>' +
          '<Cell><Data ss:Type="String">' + xesc(cr.courseName) + '</Data></Cell>' +
          '<Cell><Data ss:Type="String">' + xesc(cr.courseCode || '') + '</Data></Cell>' +
          '<Cell><Data ss:Type="Number">' + (cr.error ? 0 : cr.totalStudents) + '</Data></Cell>' +
          '<Cell><Data ss:Type="Number">' + (cr.error ? 0 : cr.atRisk.length) + '</Data></Cell>' +
        '</Row>';
      });
      xml += '</Table></Worksheet>';

      courseReports.forEach(function(cr){
        var sheetName = sanitizeSheetName(cr.courseCode || cr.courseName, usedNames);
        xml += '<Worksheet ss:Name="' + sheetName + '"><Table>' +
          '<Row>' +
            '<Cell ss:StyleID="Header"><Data ss:Type="String">Name</Data></Cell>' +
            '<Cell ss:StyleID="Header"><Data ss:Type="String">Email</Data></Cell>' +
            '<Cell ss:StyleID="Header"><Data ss:Type="String">Section</Data></Cell>' +
            '<Cell ss:StyleID="Header"><Data ss:Type="String">Last Access</Data></Cell>' +
            '<Cell ss:StyleID="Header"><Data ss:Type="String">Current Grade</Data></Cell>' +
            '<Cell ss:StyleID="Header"><Data ss:Type="String">Missing</Data></Cell>' +
            '<Cell ss:StyleID="Header"><Data ss:Type="String">Flags</Data></Cell>' +
          '</Row>';
        (cr.atRisk || []).forEach(function(s){
          xml += '<Row>' +
            '<Cell><Data ss:Type="String">' + xesc(s.name) + '</Data></Cell>' +
            '<Cell><Data ss:Type="String">' + xesc(s.email || '') + '</Data></Cell>' +
            '<Cell><Data ss:Type="String">' + xesc(s.section || '') + '</Data></Cell>' +
            '<Cell><Data ss:Type="String">' + xesc(s.lastAccessDisplay) + '</Data></Cell>' +
            '<Cell><Data ss:Type="String">' + xesc(s.gradeDisplay) + '</Data></Cell>' +
            '<Cell><Data ss:Type="String">' + xesc(s.missingDisplay) + '</Data></Cell>' +
            '<Cell><Data ss:Type="String">' + xesc(s.reasons.join('; ')) + '</Data></Cell>' +
          '</Row>';
        });
        xml += '</Table></Worksheet>';
      });
      xml += '</Workbook>';

      var blob = new Blob([xml], {type: 'application/vnd.ms-excel'});
      var url = URL.createObjectURL(blob);
      var a = document.createElement('a');
      a.href = url;
      a.download = 'at-risk-' + termName.replace(/[^a-z0-9]+/gi, '-').toLowerCase() + '.xls';
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      setTimeout(function(){ URL.revokeObjectURL(url); }, 1000);
    };
  }

  function evaluateCourse(course){
    return Promise.all([
      fetchJSON(base + '/api/v1/courses/' + course.id + '/users?enrollment_type[]=student&include[]=email&include[]=enrollments&per_page=100'),
      fetchJSON(base + '/api/v1/courses/' + course.id + '/sections?per_page=100')
    ]).then(function(res){
      var users = res[0], sections = res[1];
      var sectionMap = {};
      sections.forEach(function(sec){ sectionMap[sec.id] = sec.name; });

      return fetchJSON(base + '/api/v1/courses/' + course.id + '/analytics/student_summaries?per_page=100')
        .then(function(summaries){
          var missingMap = {};
          summaries.forEach(function(s){
            missingMap[s.id] = s.tardiness_breakdown ? s.tardiness_breakdown.missing : null;
          });
          return buildReport(course, users, sectionMap, missingMap, false);
        })
        .catch(function(){
          return buildReport(course, users, sectionMap, {}, true);
        });
    }).catch(function(err){
      return { courseName: course.name, courseCode: course.course_code, error: err.message };
    });
  }

  function buildReport(course, users, sectionMap, missingMap, missingUnavailable){
    var atRisk = [];
    users.forEach(function(u){
      var enrollment = (u.enrollments || []).filter(function(e){ return e.type === 'StudentEnrollment'; })[0] || (u.enrollments || [])[0];
      var sectionName = (enrollment && sectionMap[enrollment.course_section_id]) || '';
      var lastActivity = enrollment && enrollment.last_activity_at;
      var days = daysSince(lastActivity);
      var score = enrollment && enrollment.grades ? enrollment.grades.current_score : null;
      var missing = missingUnavailable ? null : missingMap[u.id];

      var reasons = [];
      if(days === null) reasons.push('Never accessed');
      else if(days >= INACTIVE_DAYS) reasons.push(days + ' days inactive');
      if(score != null && score < GRADE_THRESHOLD) reasons.push('Grade ' + score + '%');
      if(missing != null && missing >= MISSING_THRESHOLD) reasons.push(missing + ' missing assignments');

      if(reasons.length){
        atRisk.push({
          name: u.sortable_name || u.name,
          email: u.email,
          section: sectionName,
          lastAccessDisplay: formatDateTime(lastActivity),
          gradeDisplay: score != null ? score + '%' : 'N/A',
          missingDisplay: missing != null ? String(missing) : 'N/A',
          reasons: reasons
        });
      }
    });
    atRisk.sort(function(a, b){ return b.reasons.length - a.reasons.length; });
    return {
      courseName: course.name,
      courseCode: course.course_code,
      totalStudents: users.length,
      atRisk: atRisk,
      missingUnavailable: missingUnavailable
    };
  }

  function loadCourses(termName, matchingCourses){
    var overlay = baseOverlay('arb-overlay');
    overlay.innerHTML = '<p>Scanning ' + matchingCourses.length + ' course(s) in ' + esc(termName) + '&hellip;</p>';
    Promise.all(matchingCourses.map(evaluateCourse)).then(function(courseReports){
      buildXlsAndOverlay(termName, courseReports);
    });
  }

  function showTermPicker(courses){
    var termNames = {};
    courses.forEach(function(c){
      if(c.term && c.term.name) termNames[c.term.name] = true;
    });
    var names = Object.keys(termNames).sort();
    if(!names.length){
      alert('No terms found on your courses.');
      return;
    }
    var overlay = baseOverlay('arb-overlay');
    var html = '<h2>Pick a term</h2><div>';
    names.forEach(function(n, i){
      html += '<button class="arb-term-btn" data-idx="' + i + '" style="display:block;width:100%;text-align:left;padding:10px 14px;margin-bottom:8px;cursor:pointer;border:1px solid #999;border-radius:4px;background:#f5f5f5;">' + esc(n) + '</button>';
    });
    html += '</div><button id="arb-close-btn" style="margin-top:10px;padding:8px 14px;cursor:pointer;border:1px solid #999;border-radius:4px;background:#f5f5f5;">Cancel</button>';
    overlay.innerHTML = html;
    document.getElementById('arb-close-btn').onclick = function(){ overlay.remove(); };
    var buttons = overlay.querySelectorAll('.arb-term-btn');
    buttons.forEach(function(btn){
      btn.onclick = function(){
        var termName = names[parseInt(btn.getAttribute('data-idx'), 10)];
        var matching = courses.filter(function(c){ return c.term && c.term.name === termName; });
        loadCourses(termName, matching);
      };
    });
  }

  fetchJSON(base + '/api/v1/users/self/courses?enrollment_type=teacher&include[]=term&include[]=total_students&per_page=100&state[]=unpublished&state[]=available&state[]=completed')
    .then(function(courses){
      showTermPicker(courses);
    }).catch(function(err){
      alert('Could not load your courses: ' + err.message);
    });
}

function runPeerComparison(){
  var m = window.location.pathname.match(/\/courses\/(\d+)/);
  if(!m){ alert('Run this bookmarklet from inside a Canvas course (a URL like /courses/12345).'); return; }
  var courseId = m[1];
  var base = window.location.origin;

  function esc(s){
    return String(s == null ? '' : s).replace(/[&<>]/g, function(c){ return {'&':'&amp;','<':'&lt;','>':'&gt;'}[c]; });
  }
  function fetchJSON(url){
    return fetch(url, {credentials:'same-origin', headers:{'Accept':'application/json'}}).then(function(r){
      if(!r.ok) throw new Error('Canvas API request failed (' + r.status + '): ' + url);
      return r.json();
    });
  }
  function removeOverlay(id){
    var old = document.getElementById(id);
    if(old) old.remove();
  }
  function baseOverlay(id){
    removeOverlay(id);
    var overlay = document.createElement('div');
    overlay.id = id;
    overlay.style.cssText = 'position:fixed;top:16px;left:16px;right:16px;bottom:16px;background:#fff;z-index:999999;overflow:auto;padding:24px;border:1px solid #ccc;border-radius:6px;box-shadow:0 4px 24px rgba(0,0,0,0.35);font-family:-apple-system,Helvetica,Arial,sans-serif;color:#222;';
    document.body.appendChild(overlay);
    return overlay;
  }
  function stats(values){
    var valid = values.filter(function(v){ return v != null && !isNaN(v); });
    if(!valid.length) return null;
    var min = Math.min.apply(null, valid);
    var max = Math.max.apply(null, valid);
    var sum = valid.reduce(function(a,b){ return a+b; }, 0);
    return { min: min, max: max, avg: sum / valid.length, count: valid.length };
  }
  function percentileOf(value, others){
    var valid = others.filter(function(v){ return v != null && !isNaN(v); });
    if(value == null || isNaN(value) || !valid.length) return null;
    var belowOrEqual = 0;
    valid.forEach(function(v){ if(v <= value) belowOrEqual++; });
    return Math.round((belowOrEqual / valid.length) * 100);
  }
  function barHtml(label, valueDisplay, value, others, decimals){
    var s = stats(others);
    var pct = percentileOf(value, others);
    var barInner = '';
    if(s && value != null && !isNaN(value)){
      var range = s.max - s.min;
      var studentPos = range > 0 ? Math.max(0, Math.min(100, ((value - s.min) / range) * 100)) : 50;
      var avgPos = range > 0 ? Math.max(0, Math.min(100, ((s.avg - s.min) / range) * 100)) : 50;
      barInner = '<div style="position:relative;height:28px;background:#eee;border-radius:4px;margin:6px 0 2px;">' +
        '<div title="Class average (rest of course)" style="position:absolute;left:' + avgPos + '%;top:0;bottom:0;width:2px;background:#888;"></div>' +
        '<div title="This student" style="position:absolute;left:calc(' + studentPos + '% - 5px);top:-3px;width:10px;height:34px;background:#2e5f8a;border-radius:2px;"></div>' +
        '</div>' +
        '<div style="font-size:11px;color:#888;display:flex;justify-content:space-between;">' +
          '<span>Low (' + s.min.toFixed(decimals) + ')</span><span>High (' + s.max.toFixed(decimals) + ')</span>' +
        '</div>';
    } else {
      barInner = '<p style="color:#999;font-size:13px;"><em>Not enough classmate data to compare.</em></p>';
    }
    var pctText = pct == null ? '' : (' \u2014 higher than ' + pct + '% of the rest of the class');
    return '<div style="margin-bottom:18px;">' +
      '<div style="font-weight:600;">' + esc(label) + ': ' + esc(valueDisplay) + esc(pctText) + '</div>' +
      barInner +
      '</div>';
  }

  function renderComparison(student, others){
    var overlay = baseOverlay('pcs-overlay');
    var otherScores = others.map(function(s){ return s.score; });
    var otherPageViews = others.map(function(s){ return s.pageViews; });
    var otherParticipations = others.map(function(s){ return s.participations; });
    var otherOnTimeRates = others.map(function(s){ return s.onTimeRate; });

    var html = '<div style="text-align:right;margin-bottom:14px;">' +
      '<button id="pcs-back-btn" style="padding:8px 14px;margin-right:8px;cursor:pointer;border:1px solid #999;border-radius:4px;background:#f5f5f5;">Pick Another Student</button>' +
      '<button id="pcs-close-btn" style="padding:8px 14px;cursor:pointer;border:1px solid #999;border-radius:4px;background:#f5f5f5;">Close</button>' +
      '</div>';
    html += '<h2>Peer Comparison &mdash; ' + esc(student.name) + '</h2>' +
      '<p style="color:#666;font-size:13px;">Compared anonymously against the rest of the class. No other student\u2019s individual data is shown.</p>';

    html += barHtml('Current Grade', student.score != null ? student.score + '%' : 'N/A', student.score, otherScores, 1);
    html += barHtml('Page Views', String(student.pageViews != null ? student.pageViews : 'N/A'), student.pageViews, otherPageViews, 0);
    html += barHtml('Participations', String(student.participations != null ? student.participations : 'N/A'), student.participations, otherParticipations, 0);
    html += barHtml('On-Time Submission Rate', student.onTimeRate != null ? student.onTimeRate + '%' : 'N/A', student.onTimeRate, otherOnTimeRates, 1);

    overlay.innerHTML = html;
    document.getElementById('pcs-close-btn').onclick = function(){ overlay.remove(); };
    document.getElementById('pcs-back-btn').onclick = function(){ showPicker(student.allStudents); };
  }

  function showPicker(allStudents){
    var overlay = baseOverlay('pcs-overlay');
    var sorted = allStudents.slice().sort(function(a, b){ return (a.name || '').localeCompare(b.name || ''); });
    var html = '<h2>Peer Comparison Snapshot</h2>' +
      '<p style="color:#666;font-size:13px;">Pick a student to see how they compare to the rest of the class &mdash; anonymously.</p>' +
      '<select id="pcs-select" style="width:100%;padding:8px;margin-bottom:12px;font-size:14px;">';
    sorted.forEach(function(s){
      html += '<option value="' + esc(s.id) + '">' + esc(s.name) + '</option>';
    });
    html += '</select>' +
      '<button id="pcs-go-btn" style="padding:8px 14px;cursor:pointer;border:1px solid #999;border-radius:4px;background:#f5f5f5;">Show Comparison</button> ' +
      '<button id="pcs-close-btn" style="padding:8px 14px;cursor:pointer;border:1px solid #999;border-radius:4px;background:#f5f5f5;">Cancel</button>';
    overlay.innerHTML = html;
    document.getElementById('pcs-close-btn').onclick = function(){ overlay.remove(); };
    document.getElementById('pcs-go-btn').onclick = function(){
      var selectedId = document.getElementById('pcs-select').value;
      var student = null, others = [];
      allStudents.forEach(function(s){
        if(String(s.id) === String(selectedId)) student = s;
        else others.push(s);
      });
      student.allStudents = allStudents;
      renderComparison(student, others);
    };
  }

  Promise.all([
    fetchJSON(base + '/api/v1/courses/' + courseId + '/users?enrollment_type[]=student&include[]=enrollments&per_page=100'),
    fetchJSON(base + '/api/v1/courses/' + courseId + '/analytics/student_summaries?per_page=100')
  ]).then(function(res){
    var users = res[0], summaries = res[1];
    var summaryMap = {};
    summaries.forEach(function(s){ summaryMap[s.id] = s; });

    var allStudents = users.map(function(u){
      var enrollment = (u.enrollments || []).filter(function(e){ return e.type === 'StudentEnrollment'; })[0] || (u.enrollments || [])[0];
      var score = enrollment && enrollment.grades ? enrollment.grades.current_score : null;
      var summary = summaryMap[u.id];
      var tb = summary && summary.tardiness_breakdown;
      var onTimeRate = null;
      if(tb){
        var denom = (tb.on_time || 0) + (tb.late || 0) + (tb.missing || 0);
        if(denom > 0) onTimeRate = Math.round((tb.on_time / denom) * 1000) / 10;
      }
      return {
        id: u.id,
        name: u.sortable_name || u.name,
        score: score,
        pageViews: summary ? summary.page_views : null,
        participations: summary ? summary.participations : null,
        onTimeRate: onTimeRate
      };
    });

    showPicker(allStudents);
  }).catch(function(err){
    alert('Could not load comparison data: ' + err.message);
  });
}

function runReengagementDrafter(){
  var base = window.location.origin;
  var INACTIVE_DAYS = 7;
  var GRADE_THRESHOLD = 70;
  var MISSING_THRESHOLD = 2;
  var MAX_SUBMISSION_PAGES = 5;

  function esc(s){
    return (s || '').replace(/[&<>]/g, function(c){ return {'&':'&amp;','<':'&lt;','>':'&gt;'}[c]; });
  }
  function fetchJSON(url){
    return fetch(url, {credentials:'same-origin', headers:{'Accept':'application/json'}}).then(function(r){
      if(!r.ok) throw new Error('Canvas API request failed (' + r.status + '): ' + url);
      return r.json();
    });
  }
  function parseNextLink(header){
    if(!header) return null;
    var parts = header.split(',');
    for(var i=0;i<parts.length;i++){
      var m = parts[i].match(/<([^>]+)>;\s*rel="next"/);
      if(m) return m[1];
    }
    return null;
  }
  function fetchPaged(url, maxPages){
    var results = [];
    var pages = 0;
    function step(u){
      return fetch(u, {credentials:'same-origin', headers:{'Accept':'application/json'}}).then(function(r){
        if(!r.ok) throw new Error('Canvas API request failed (' + r.status + '): ' + u);
        pages++;
        return r.json().then(function(data){
          results = results.concat(data);
          var next = parseNextLink(r.headers.get('Link'));
          if(next && pages < maxPages) return step(next);
          return results;
        });
      });
    }
    return step(url);
  }
  function removeOverlay(id){
    var old = document.getElementById(id);
    if(old) old.remove();
  }
  function baseOverlay(id){
    removeOverlay(id);
    var overlay = document.createElement('div');
    overlay.id = id;
    overlay.style.cssText = 'position:fixed;top:16px;left:16px;right:16px;bottom:16px;background:#fff;z-index:999999;overflow:auto;padding:24px;border:1px solid #ccc;border-radius:6px;box-shadow:0 4px 24px rgba(0,0,0,0.35);font-family:-apple-system,Helvetica,Arial,sans-serif;color:#222;';
    document.body.appendChild(overlay);
    return overlay;
  }
  function daysSince(dateStr){
    if(!dateStr) return null;
    var diffMs = Date.now() - new Date(dateStr).getTime();
    return Math.floor(diffMs / (1000 * 60 * 60 * 24));
  }
  function formatDateTime(d){
    if(!d) return 'Never';
    var dt = new Date(d);
    return dt.toLocaleDateString('en-US', {month:'short', day:'numeric', year:'numeric'}) +
      ' ' + dt.toLocaleTimeString('en-US', {hour:'numeric', minute:'2-digit'});
  }
  function firstName(sortableName){
    if(!sortableName) return 'there';
    if(sortableName.indexOf(',') !== -1){
      var afterComma = sortableName.split(',')[1] || '';
      var trimmed = afterComma.trim().split(' ')[0];
      return trimmed || sortableName;
    }
    return sortableName.split(' ')[0];
  }
  function copyText(text){
    if(navigator.clipboard && navigator.clipboard.writeText){
      navigator.clipboard.writeText(text);
    } else {
      var ta = document.createElement('textarea');
      ta.value = text;
      document.body.appendChild(ta);
      ta.select();
      document.execCommand('copy');
      document.body.removeChild(ta);
    }
  }

  function buildDraft(instructorName, courseName, student){
    var subject = 'Checking in \u2014 ' + courseName;
    var observations = [];

    if(student.lastAccessDisplay === 'Never'){
      observations.push("I noticed you haven't logged into our Canvas course yet this term.");
    } else if(student.daysInactive !== null && student.daysInactive >= INACTIVE_DAYS){
      observations.push('I noticed you haven\u2019t been active in our Canvas course since ' + student.lastAccessDisplay + '.');
    }

    if(student.missingNames && student.missingNames.length){
      observations.push('You currently have the following assignment(s) missing: ' + student.missingNames.join(', ') + '.');
    } else if(student.missingCount != null && student.missingCount >= MISSING_THRESHOLD){
      observations.push('You currently have ' + student.missingCount + ' missing assignments in the course.');
    }

    if(student.score != null && student.score < GRADE_THRESHOLD){
      observations.push('Your current grade in the course is ' + student.score + '%.');
    }

    var body = 'Hi ' + firstName(student.name) + ',\n\n' +
      observations.join(' ') + '\n\n' +
      'I wanted to check in and see how things are going. If you\u2019re running into any challenges, I\u2019m happy to help \u2014 feel free to reply to this email or stop by office hours.\n\n' +
      'Best,\n' + instructorName;

    return { subject: subject, body: body };
  }

  function buildReport(course, users, sectionMap, missingCountMap, missingCountUnavailable, missingNamesMap, missingNamesUnavailable){
    var atRisk = [];
    users.forEach(function(u){
      var enrollment = (u.enrollments || []).filter(function(e){ return e.type === 'StudentEnrollment'; })[0] || (u.enrollments || [])[0];
      var lastActivity = enrollment && enrollment.last_activity_at;
      var days = daysSince(lastActivity);
      var score = enrollment && enrollment.grades ? enrollment.grades.current_score : null;
      var missingCount = missingCountUnavailable ? null : missingCountMap[u.id];

      var flagged = false;
      if(days === null) flagged = true;
      else if(days >= INACTIVE_DAYS) flagged = true;
      if(score != null && score < GRADE_THRESHOLD) flagged = true;
      if(missingCount != null && missingCount >= MISSING_THRESHOLD) flagged = true;

      if(flagged){
        atRisk.push({
          id: u.id,
          name: u.sortable_name || u.name,
          email: u.email,
          lastAccessDisplay: formatDateTime(lastActivity),
          daysInactive: days,
          score: score,
          missingCount: missingCount,
          missingNames: missingNamesUnavailable ? null : (missingNamesMap[u.id] || [])
        });
      }
    });
    return {
      courseId: course.id,
      courseName: course.name,
      courseCode: course.course_code,
      atRisk: atRisk,
      missingNamesUnavailable: missingNamesUnavailable
    };
  }

  function evaluateCourse(course){
    var usersPromise = fetchJSON(base + '/api/v1/courses/' + course.id + '/users?enrollment_type[]=student&include[]=email&include[]=enrollments&per_page=100');
    var sectionsPromise = fetchJSON(base + '/api/v1/courses/' + course.id + '/sections?per_page=100');
    var summariesPromise = fetchJSON(base + '/api/v1/courses/' + course.id + '/analytics/student_summaries?per_page=100').catch(function(){ return null; });
    var submissionsPromise = fetchPaged(base + '/api/v1/courses/' + course.id + '/students/submissions?student_ids[]=all&per_page=100&include[]=assignment', MAX_SUBMISSION_PAGES).catch(function(){ return null; });

    return Promise.all([usersPromise, sectionsPromise, summariesPromise, submissionsPromise]).then(function(res){
      var users = res[0], sections = res[1], summaries = res[2], submissions = res[3];
      var sectionMap = {};
      sections.forEach(function(sec){ sectionMap[sec.id] = sec.name; });

      var missingCountMap = {}, missingCountUnavailable = !summaries;
      if(summaries){
        summaries.forEach(function(s){
          missingCountMap[s.id] = s.tardiness_breakdown ? s.tardiness_breakdown.missing : null;
        });
      }

      var missingNamesMap = {}, missingNamesUnavailable = !submissions;
      if(submissions){
        var now = new Date();
        submissions.forEach(function(sub){
          if(sub.missing){
            if(sub.cached_due_date && new Date(sub.cached_due_date) > now) return;
            var name = (sub.assignment && sub.assignment.name) || ('Assignment ' + sub.assignment_id);
            if(!missingNamesMap[sub.user_id]) missingNamesMap[sub.user_id] = [];
            if(missingNamesMap[sub.user_id].indexOf(name) === -1) missingNamesMap[sub.user_id].push(name);
          }
        });
      }

      return buildReport(course, users, sectionMap, missingCountMap, missingCountUnavailable, missingNamesMap, missingNamesUnavailable);
    }).catch(function(err){
      return { courseName: course.name, courseCode: course.course_code, error: err.message, atRisk: [] };
    });
  }

  function messageHref(courseId, uid, name){
    return base + '/conversations?context_id=course_' + courseId + '&user_id=' + encodeURIComponent(uid) + '&user_name=' + encodeURIComponent(name);
  }
  function legacyCopy(ta){
    try {
      ta.focus(); ta.select();
      return document.execCommand('copy');
    } catch(e){ return false; }
  }
  function setNativeValue(win, el, val){
    var proto = el.tagName === 'TEXTAREA' ? win.HTMLTextAreaElement.prototype : win.HTMLInputElement.prototype;
    var setter = Object.getOwnPropertyDescriptor(proto, 'value').set;
    setter.call(el, val);
    el.dispatchEvent(new win.Event('input', {bubbles: true}));
    el.dispatchEvent(new win.Event('change', {bubbles: true}));
  }
  function tryAutoFillMessage(win, subjectText, bodyText, onDone){
    var attempts = 0;
    var maxAttempts = 40;
    var timer = setInterval(function(){
      attempts++;
      if(win.closed){ clearInterval(timer); onDone(false, false); return; }
      var textarea = null, subjectInput = null;
      try {
        var doc = win.document;
        textarea = doc.querySelector('textarea[data-testid="message-body"]') ||
          doc.querySelector('textarea[name="body"]');
        if(!textarea){
          var areas = Array.prototype.slice.call(doc.querySelectorAll('textarea'));
          var visible = areas.filter(function(t){ return t.offsetWidth > 0 && t.offsetHeight > 0; });
          if(visible.length){
            visible.sort(function(a, b){ return (b.offsetWidth * b.offsetHeight) - (a.offsetWidth * a.offsetHeight); });
            textarea = visible[0];
          }
        }
        subjectInput = doc.querySelector('input[data-testid="subject-input"]') ||
          doc.querySelector('input[name="subject"]') ||
          doc.querySelector('input[placeholder="Subject"]');
      } catch(e){
        textarea = null;
      }
      if(textarea){
        clearInterval(timer);
        setNativeValue(win, textarea, bodyText);
        if(subjectInput) setNativeValue(win, subjectInput, subjectText);
        onDone(true, !!subjectInput);
        return;
      }
      if(attempts >= maxAttempts){
        clearInterval(timer);
        onDone(false, false);
      }
    }, 200);
  }

  function renderDrafts(instructorName, termName, courseReports){
    var overlay = baseOverlay('red-overlay');
    var html = '<div style="text-align:right;margin-bottom:14px;position:sticky;top:0;background:#fff;padding-bottom:8px;border-bottom:1px solid #eee;">' +
      '<button id="red-close-btn" style="padding:8px 14px;cursor:pointer;border:1px solid #999;border-radius:4px;background:#f5f5f5;">Close</button>' +
      '</div>';
    html += '<h2>' + esc(termName) + ' &mdash; Re-Engagement Drafts</h2>' +
      '<p style="color:#666;font-size:13px;">Every draft is editable before you send it. Nothing is sent automatically.</p>';

    var cardIndex = 0;
    var cardData = [];

    courseReports.forEach(function(cr){
      html += '<h3>' + esc(cr.courseName) + (cr.courseCode ? ' (' + esc(cr.courseCode) + ')' : '') + '</h3>';
      if(cr.error){
        html += '<p><em>Could not load data: ' + esc(cr.error) + '</em></p>';
        return;
      }
      if(cr.missingNamesUnavailable){
        html += '<p style="color:#a66;font-size:13px;"><em>Could not pull specific missing-assignment names for this course &mdash; drafts will reference the missing count instead where applicable.</em></p>';
      }
      if(!cr.atRisk.length){
        html += '<p><em>No students currently flagged.</em></p>';
        return;
      }
      cr.atRisk.forEach(function(student){
        var draft = buildDraft(instructorName, cr.courseName, student);
        var idx = cardIndex++;
        cardData.push({ email: student.email, uid: student.id, courseId: cr.courseId, name: student.name });
        html += '<div style="border:1px solid #ddd;border-radius:6px;padding:16px;margin-bottom:16px;">' +
          '<div style="font-weight:600;margin-bottom:4px;">' + esc(student.name) + ' &lt;' + esc(student.email || '') + '&gt;</div>' +
          '<div style="font-size:13px;color:#666;margin-bottom:10px;">Last access: ' + esc(student.lastAccessDisplay) +
            (student.score != null ? ' &middot; Grade: ' + student.score + '%' : '') +
            (student.missingCount != null ? ' &middot; Missing: ' + student.missingCount : '') + '</div>' +
          'Subject:<br><input type="text" id="red-subject-' + idx + '" value="' + esc(draft.subject) + '" style="width:100%;padding:6px;margin-bottom:8px;box-sizing:border-box;">' +
          'Body:<br><textarea id="red-body-' + idx + '" rows="8" style="width:100%;padding:6px;box-sizing:border-box;font-family:inherit;">' + esc(draft.body) + '</textarea>' +
          '<div style="margin-top:8px;">' +
            '<button class="red-mail-btn" data-idx="' + idx + '" style="padding:6px 12px;margin-right:8px;cursor:pointer;border:1px solid #999;border-radius:4px;background:#f5f5f5;">Open in Canvas Message</button>' +
            '<button class="red-copy-btn" data-idx="' + idx + '" style="padding:6px 12px;cursor:pointer;border:1px solid #999;border-radius:4px;background:#f5f5f5;">Copy Body</button>' +
          '</div>' +
        '</div>';
      });
    });

    overlay.innerHTML = html;
    document.getElementById('red-close-btn').onclick = function(){ overlay.remove(); };

    var mailBtns = overlay.querySelectorAll('.red-mail-btn');
    mailBtns.forEach(function(btn){
      btn.onclick = function(){
        var idx = btn.getAttribute('data-idx');
        var card = cardData[idx];
        var subjectField = document.getElementById('red-subject-' + idx);
        var bodyField = document.getElementById('red-body-' + idx);
        var subjectText = subjectField.value;
        var bodyText = bodyField.value;

        var win = window.open(messageHref(card.courseId, card.uid, card.name), '_blank');

        function clipboardFallback(){
          function done(ok){
            alert(ok ?
              'Opened the Canvas message and copied the draft to your clipboard \u2014 paste it into the message body.' :
              'Opened the Canvas message. Could not copy automatically \u2014 select the text in the box above and copy it manually (Ctrl/Cmd+C), then paste into the message.');
          }
          if(navigator.clipboard && navigator.clipboard.writeText){
            navigator.clipboard.writeText(bodyText).then(function(){ done(true); }, function(){ done(legacyCopy(bodyField)); });
          } else {
            done(legacyCopy(bodyField));
          }
        }

        if(win){
          tryAutoFillMessage(win, subjectText, bodyText, function(filled, subjectFilled){
            if(filled){
              alert('The message body' + (subjectFilled ? ' and subject have' : ' has') + ' been filled in \u2014 review it in the Canvas tab that just opened, then click Send.');
            } else {
              clipboardFallback();
            }
          });
        } else {
          clipboardFallback();
        }
      };
    });
    var copyBtns = overlay.querySelectorAll('.red-copy-btn');
    copyBtns.forEach(function(btn){
      btn.onclick = function(){
        var idx = btn.getAttribute('data-idx');
        copyText(document.getElementById('red-body-' + idx).value);
      };
    });
  }

  function loadCourses(instructorName, termName, matchingCourses){
    var overlay = baseOverlay('red-overlay');
    overlay.innerHTML = '<p>Scanning ' + matchingCourses.length + ' course(s) in ' + esc(termName) + ' and drafting messages&hellip;</p>';
    Promise.all(matchingCourses.map(evaluateCourse)).then(function(courseReports){
      renderDrafts(instructorName, termName, courseReports);
    });
  }

  function showTermPicker(instructorName, courses){
    var termNames = {};
    courses.forEach(function(c){
      if(c.term && c.term.name) termNames[c.term.name] = true;
    });
    var names = Object.keys(termNames).sort();
    if(!names.length){
      alert('No terms found on your courses.');
      return;
    }
    var overlay = baseOverlay('red-overlay');
    var html = '<h2>Pick a term</h2><div>';
    names.forEach(function(n, i){
      html += '<button class="red-term-btn" data-idx="' + i + '" style="display:block;width:100%;text-align:left;padding:10px 14px;margin-bottom:8px;cursor:pointer;border:1px solid #999;border-radius:4px;background:#f5f5f5;">' + esc(n) + '</button>';
    });
    html += '</div><button id="red-close-btn" style="margin-top:10px;padding:8px 14px;cursor:pointer;border:1px solid #999;border-radius:4px;background:#f5f5f5;">Cancel</button>';
    overlay.innerHTML = html;
    document.getElementById('red-close-btn').onclick = function(){ overlay.remove(); };
    var buttons = overlay.querySelectorAll('.red-term-btn');
    buttons.forEach(function(btn){
      btn.onclick = function(){
        var termName = names[parseInt(btn.getAttribute('data-idx'), 10)];
        var matching = courses.filter(function(c){ return c.term && c.term.name === termName; });
        loadCourses(instructorName, termName, matching);
      };
    });
  }

  Promise.all([
    fetchJSON(base + '/api/v1/users/self'),
    fetchJSON(base + '/api/v1/users/self/courses?enrollment_type=teacher&include[]=term&per_page=100&state[]=unpublished&state[]=available&state[]=completed')
  ]).then(function(res){
    var me = res[0], courses = res[1];
    showTermPicker(me.name || me.short_name || 'Your Instructor', courses);
  }).catch(function(err){
    alert('Could not load your courses: ' + err.message);
  });
}

function runWeeklyAnnouncementDrafter(){
  var m = window.location.pathname.match(/\/courses\/(\d+)/);
  if(!m){ alert('Run this bookmarklet from inside a Canvas course (a URL like /courses/12345).'); return; }
  var courseId = m[1];
  var base = window.location.origin;

  function esc(s){
    return String(s == null ? '' : s).replace(/[&<>]/g, function(c){ return {'&':'&amp;','<':'&lt;','>':'&gt;'}[c]; });
  }
  function fetchJSON(url){
    return fetch(url, {credentials:'same-origin', headers:{'Accept':'application/json'}}).then(function(r){
      if(!r.ok) throw new Error('Canvas API request failed (' + r.status + '): ' + url);
      return r.json();
    });
  }
  function mondayOf(d){
    var dt = new Date(d);
    var day = dt.getDay();
    var diff = (day === 0 ? -6 : 1 - day);
    var monday = new Date(dt);
    monday.setDate(dt.getDate() + diff);
    monday.setHours(0,0,0,0);
    return monday;
  }
  function formatDate(d){
    var dt = new Date(d);
    return dt.toLocaleDateString('en-US', {weekday:'short', month:'short', day:'numeric'});
  }
  function weekRangeLabel(monday, sunday){
    var opts = {month:'short', day:'numeric'};
    return monday.toLocaleDateString('en-US', opts) + ' \u2013 ' + sunday.toLocaleDateString('en-US', opts);
  }
  function tryAutoFillAnnouncement(win, bodyHtml, onDone){
    var attempts = 0;
    var maxAttempts = 50;
    var timer = setInterval(function(){
      attempts++;
      if(win.closed){ clearInterval(timer); onDone(false); return; }
      var iframe = null;
      try {
        var doc = win.document;
        iframe = doc.querySelector('iframe.tox-edit-area__iframe') ||
          doc.querySelector('iframe[id^="announcement_message_ifr"]') ||
          doc.querySelector('iframe[title="Rich Content Editor"]') ||
          doc.querySelector('.rce-wrapper iframe');
      } catch(e){
        iframe = null;
      }
      if(iframe && iframe.contentDocument && iframe.contentDocument.body){
        clearInterval(timer);
        var iwin = iframe.contentWindow;
        var body = iframe.contentDocument.body;
        body.innerHTML = bodyHtml;
        body.dispatchEvent(new iwin.Event('input', {bubbles: true}));
        body.dispatchEvent(new iwin.Event('keyup', {bubbles: true}));
        body.dispatchEvent(new iwin.Event('change', {bubbles: true}));
        onDone(true);
        return;
      }
      if(attempts >= maxAttempts){
        clearInterval(timer);
        onDone(false);
      }
    }, 200);
  }
  function removeOverlay(id){
    var old = document.getElementById(id);
    if(old) old.remove();
  }
  function baseOverlay(id){
    removeOverlay(id);
    var overlay = document.createElement('div');
    overlay.id = id;
    overlay.style.cssText = 'position:fixed;top:16px;left:16px;right:16px;bottom:16px;background:#fff;z-index:999999;overflow:auto;padding:24px;border:1px solid #ccc;border-radius:6px;box-shadow:0 4px 24px rgba(0,0,0,0.35);font-family:-apple-system,Helvetica,Arial,sans-serif;color:#222;';
    document.body.appendChild(overlay);
    return overlay;
  }

  var now = new Date();
  var weekStart = mondayOf(now);
  var weekEnd = new Date(weekStart);
  weekEnd.setDate(weekStart.getDate() + 6);
  weekEnd.setHours(23, 59, 59, 999);

  Promise.all([
    fetchJSON(base + '/api/v1/courses/' + courseId),
    fetchJSON(base + '/api/v1/courses/' + courseId + '/modules?include[]=items&per_page=100'),
    fetchJSON(base + '/api/v1/courses/' + courseId + '/assignments?per_page=100')
  ]).then(function(res){
    var course = res[0], modules = res[1], assignments = res[2];
    var byAssignmentId = {}, byQuizId = {}, byDiscussionTopicId = {};
    assignments.forEach(function(a){
      if(a.published === false) return;
      byAssignmentId[a.id] = a;
      if(a.quiz_id) byQuizId[a.quiz_id] = a;
      if(a.discussion_topic && a.discussion_topic.id) byDiscussionTopicId[a.discussion_topic.id] = a;
    });

    function inWeek(dateStr){
      if(!dateStr) return false;
      var d = new Date(dateStr);
      return d >= weekStart && d <= weekEnd;
    }

    var dueThisWeek = [];
    var activeModuleIds = {};
    var newModulesThisWeek = [];

    modules.forEach(function(mod){
      if(mod.published === false) return;
      if(mod.unlock_at && inWeek(mod.unlock_at)){
        newModulesThisWeek.push(mod.name);
      }
      (mod.items || []).forEach(function(item){
        if(item.published === false) return;
        var matched = null;
        if(item.type === 'Assignment') matched = byAssignmentId[item.content_id];
        else if(item.type === 'Quiz') matched = byQuizId[item.content_id];
        else if(item.type === 'Discussion') matched = byDiscussionTopicId[item.content_id];
        if(matched && matched.due_at && inWeek(matched.due_at)){
          dueThisWeek.push({ title: item.title, due_at: new Date(matched.due_at) });
          activeModuleIds[mod.id] = true;
        }
      });
    });
    dueThisWeek.sort(function(a, b){ return a.due_at - b.due_at; });

    var resourcesThisWeek = [];
    modules.forEach(function(mod){
      if(mod.published === false) return;
      if(!activeModuleIds[mod.id]) return;
      (mod.items || []).forEach(function(item){
        if(item.published === false) return;
        if(item.type === 'Page' || item.type === 'File' || item.type === 'ExternalUrl' || item.type === 'ExternalTool'){
          resourcesThisWeek.push(item.title);
        }
      });
    });

    var weekLabel = weekRangeLabel(weekStart, weekEnd);
    var subject = 'This week in ' + course.name + ': ' + weekLabel;

    var bodyHtml = '<p>Hi everyone,</p><p>Here\u2019s what\u2019s happening this week (' + esc(weekLabel) + '):</p>';

    if(dueThisWeek.length){
      bodyHtml += '<h4>Due this week</h4><ul>';
      dueThisWeek.forEach(function(d){
        bodyHtml += '<li>' + esc(d.title) + ' \u2014 due ' + esc(formatDate(d.due_at)) + '</li>';
      });
      bodyHtml += '</ul>';
    }
    if(newModulesThisWeek.length){
      bodyHtml += '<h4>New this week</h4><ul>';
      newModulesThisWeek.forEach(function(name){
        bodyHtml += '<li>' + esc(name) + '</li>';
      });
      bodyHtml += '</ul>';
    }
    if(resourcesThisWeek.length){
      bodyHtml += '<h4>This week\u2019s resources</h4><ul>';
      resourcesThisWeek.forEach(function(title){
        bodyHtml += '<li>' + esc(title) + '</li>';
      });
      bodyHtml += '</ul>';
    }
    bodyHtml += '<p>Let me know if you have any questions!</p>';

    var overlay = baseOverlay('wad-overlay');
    var noteHtml = '';
    if(!dueThisWeek.length && !newModulesThisWeek.length && !resourcesThisWeek.length){
      noteHtml = '<p style="color:#a66;font-size:13px;"><em>No due dates or newly-unlocked modules found for this week &mdash; you may want to fill this in from scratch.</em></p>';
    }

    overlay.innerHTML =
      '<div style="text-align:right;margin-bottom:14px;position:sticky;top:0;background:#fff;padding-bottom:8px;border-bottom:1px solid #eee;">' +
        '<button id="wad-open-btn" style="padding:8px 14px;margin-right:8px;cursor:pointer;border:1px solid #999;border-radius:4px;background:#f5f5f5;">Open in Canvas Announcements</button>' +
        '<button id="wad-close-btn" style="padding:8px 14px;cursor:pointer;border:1px solid #999;border-radius:4px;background:#f5f5f5;">Close</button>' +
      '</div>' +
      '<h2>Weekly Announcement Draft</h2>' + noteHtml +
      '<label style="display:block;font-weight:600;margin-bottom:4px;">Title</label>' +
      '<input type="text" id="wad-subject" value="' + esc(subject) + '" style="width:100%;padding:8px;margin-bottom:14px;font-size:14px;box-sizing:border-box;">' +
      '<label style="display:block;font-weight:600;margin-bottom:4px;">Message (edit freely, then Open in Canvas Announcements)</label>' +
      '<div id="wad-body" contenteditable="true" style="border:1px solid #ccc;border-radius:4px;padding:12px;min-height:200px;">' + bodyHtml + '</div>' +
      '<p style="color:#666;font-size:13px;margin-top:14px;">This does not post anything. Click <strong>Open in Canvas Announcements</strong> to open a new announcement for this course with the message filled in \u2014 the title is left blank for you to fill in, and nothing posts until you click Save yourself.</p>';

    document.getElementById('wad-close-btn').onclick = function(){ overlay.remove(); };
    document.getElementById('wad-open-btn').onclick = function(){
      var bodyHtml2 = document.getElementById('wad-body').innerHTML;
      var win = window.open(base + '/courses/' + courseId + '/discussion_topics/new?is_announcement=true', '_blank');

      function clipboardFallback(){
        var range = document.createRange();
        range.selectNode(document.getElementById('wad-body'));
        var sel = window.getSelection();
        sel.removeAllRanges();
        sel.addRange(range);
        var copied = document.execCommand('copy');
        sel.removeAllRanges();
        alert(copied ?
          'Opened the New Announcement page and copied the draft to your clipboard \u2014 paste it into the message box, add a title, and click Save.' :
          'Opened the New Announcement page. Could not copy automatically \u2014 select the text above and copy it manually (Ctrl/Cmd+C), then paste it in, add a title, and click Save.');
      }

      if(win){
        tryAutoFillAnnouncement(win, bodyHtml2, function(filled){
          if(filled){
            alert('The announcement body has been filled in \u2014 add a title, review it, then click Save in the tab that just opened.');
          } else {
            clipboardFallback();
          }
        });
      } else {
        clipboardFallback();
      }
    };
  }).catch(function(err){
    alert('Could not build the announcement draft: ' + err.message);
  });
}function runDiscussionExport(){
  var m = window.location.pathname.match(/\/courses\/(\d+)\/discussion_topics\/(\d+)/);
  if(!m){ alert('Run this bookmarklet from a Canvas discussion board page (a URL like /courses/12345/discussion_topics/6789).'); return; }
  var courseId = m[1];
  var topicId = m[2];
  var base = window.location.origin;

  function fetchJSON(url){
    return fetch(url, {credentials:'same-origin', headers:{'Accept':'application/json'}}).then(function(r){
      if(!r.ok) throw new Error('Canvas API request failed (' + r.status + '): ' + url);
      return r.json();
    });
  }
  function xesc(s){
    return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  }
  function htmlToLines(html){
    if(!html) return [];
    var text = html
      .replace(/<\/(p|div|li|h[1-6]|blockquote|tr)>/gi, '\n')
      .replace(/<br\s*\/?>/gi, '\n')
      .replace(/<li[^>]*>/gi, '\u2022 ')
      .replace(/<[^>]+>/g, '');
    var ta = document.createElement('textarea');
    ta.innerHTML = text;
    text = ta.value;
    var lines = text.split('\n').map(function(l){ return l.replace(/\s+$/, ''); });
    var result = [];
    var lastBlank = true;
    lines.forEach(function(l){
      var trimmed = l.trim();
      if(!trimmed){
        if(!lastBlank) result.push('');
        lastBlank = true;
      } else {
        result.push(l);
        lastBlank = false;
      }
    });
    while(result.length && result[result.length - 1] === '') result.pop();
    return result;
  }
  function formatDate(d){
    if(!d) return '';
    var dt = new Date(d);
    return dt.toLocaleDateString('en-US', {month:'short', day:'numeric', year:'numeric'}) +
      ' ' + dt.toLocaleTimeString('en-US', {hour:'numeric', minute:'2-digit'});
  }

  function run(text, opts){
    opts = opts || {};
    var rPr = '';
    if(opts.bold || opts.italic || opts.size){
      rPr = '<w:rPr>' +
        (opts.bold ? '<w:b/>' : '') +
        (opts.italic ? '<w:i/>' : '') +
        (opts.color ? '<w:color w:val="' + opts.color + '"/>' : '') +
        (opts.size ? '<w:sz w:val="' + opts.size + '"/><w:szCs w:val="' + opts.size + '"/>' : '') +
        '</w:rPr>';
    }
    return '<w:r>' + rPr + '<w:t xml:space="preserve">' + xesc(text) + '</w:t></w:r>';
  }
  function paragraph(text, opts){
    opts = opts || {};
    var pPr = '';
    var pPrParts = [];
    if(opts.indent) pPrParts.push('<w:ind w:left="' + opts.indent + '"/>');
    if(opts.spacingAfter != null) pPrParts.push('<w:spacing w:after="' + opts.spacingAfter + '"/>');
    if(pPrParts.length) pPr = '<w:pPr>' + pPrParts.join('') + '</w:pPr>';
    return '<w:p>' + pPr + (text === '' ? '' : run(text, opts)) + '</w:p>';
  }

  function renderEntries(entries, depth, out){
    (entries || []).forEach(function(entry){
      if(entry.deleted){
        out.push(paragraph('[deleted post]', {indent: Math.min(depth, 6) * 360, italic: true, color: '999999', size: 20, spacingAfter: 120}));
      } else {
        var author = entry.user_id != null && window.__pcParticipants[entry.user_id] ? window.__pcParticipants[entry.user_id] : 'Unknown';
        out.push(paragraph(author + '  \u2014  ' + formatDate(entry.created_at), {indent: Math.min(depth, 6) * 360, bold: true, size: 20, spacingAfter: 40}));
        var lines = htmlToLines(entry.message);
        if(!lines.length) lines = ['(no text)'];
        lines.forEach(function(line, i){
          out.push(paragraph(line, {indent: Math.min(depth, 6) * 360, size: 22, spacingAfter: (i === lines.length - 1) ? 160 : 0}));
        });
      }
      if(entry.replies && entry.replies.length){
        renderEntries(entry.replies, depth + 1, out);
      }
    });
  }
  function countEntries(entries){
    var count = 0;
    (entries || []).forEach(function(e){
      count++;
      if(e.replies && e.replies.length) count += countEntries(e.replies);
    });
    return count;
  }

  function crc32Table(){
    var table = [];
    for(var n = 0; n < 256; n++){
      var c = n;
      for(var k = 0; k < 8; k++){
        c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
      }
      table[n] = c >>> 0;
    }
    return table;
  }
  var CRC_TABLE = crc32Table();
  function crc32(bytes){
    var crc = 0xFFFFFFFF;
    for(var i = 0; i < bytes.length; i++){
      crc = (crc >>> 8) ^ CRC_TABLE[(crc ^ bytes[i]) & 0xFF];
    }
    return (crc ^ 0xFFFFFFFF) >>> 0;
  }
  function makeZip(files){
    var encoder = new TextEncoder();
    var localParts = [];
    var centralParts = [];
    var offset = 0;

    files.forEach(function(file){
      var nameBytes = encoder.encode(file.name);
      var data = file.data;
      var crc = crc32(data);
      var size = data.length;

      var local = new Uint8Array(30 + nameBytes.length);
      var dv = new DataView(local.buffer);
      dv.setUint32(0, 0x04034b50, true);
      dv.setUint16(4, 20, true);
      dv.setUint16(6, 0, true);
      dv.setUint16(8, 0, true);
      dv.setUint16(10, 0, true);
      dv.setUint16(12, 0x21, true);
      dv.setUint32(14, crc, true);
      dv.setUint32(18, size, true);
      dv.setUint32(22, size, true);
      dv.setUint16(26, nameBytes.length, true);
      dv.setUint16(28, 0, true);
      local.set(nameBytes, 30);

      localParts.push(local);
      localParts.push(data);

      var central = new Uint8Array(46 + nameBytes.length);
      var cdv = new DataView(central.buffer);
      cdv.setUint32(0, 0x02014b50, true);
      cdv.setUint16(4, 20, true);
      cdv.setUint16(6, 20, true);
      cdv.setUint16(8, 0, true);
      cdv.setUint16(10, 0, true);
      cdv.setUint16(12, 0, true);
      cdv.setUint16(14, 0x21, true);
      cdv.setUint32(16, crc, true);
      cdv.setUint32(20, size, true);
      cdv.setUint32(24, size, true);
      cdv.setUint16(28, nameBytes.length, true);
      cdv.setUint16(30, 0, true);
      cdv.setUint16(32, 0, true);
      cdv.setUint16(34, 0, true);
      cdv.setUint16(36, 0, true);
      cdv.setUint32(38, 0, true);
      cdv.setUint32(42, offset, true);
      central.set(nameBytes, 46);

      centralParts.push(central);
      offset += local.length + data.length;
    });

    var centralStart = offset;
    var centralSize = centralParts.reduce(function(sum, p){ return sum + p.length; }, 0);

    var eocd = new Uint8Array(22);
    var edv = new DataView(eocd.buffer);
    edv.setUint32(0, 0x06054b50, true);
    edv.setUint16(4, 0, true);
    edv.setUint16(6, 0, true);
    edv.setUint16(8, files.length, true);
    edv.setUint16(10, files.length, true);
    edv.setUint32(12, centralSize, true);
    edv.setUint32(16, centralStart, true);
    edv.setUint16(20, 0, true);

    var allParts = localParts.concat(centralParts).concat([eocd]);
    return new Blob(allParts, {type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'});
  }

  Promise.all([
    fetchJSON(base + '/api/v1/courses/' + courseId + '/discussion_topics/' + topicId),
    fetchJSON(base + '/api/v1/courses/' + courseId + '/discussion_topics/' + topicId + '/view')
  ]).then(function(res){
    var topic = res[0], view = res[1];
    var participants = {};
    (view.participants || []).forEach(function(p){ participants[p.id] = p.display_name; });
    window.__pcParticipants = participants;

    var body = [];
    body.push(paragraph(topic.title || 'Discussion', {bold: true, size: 32, spacingAfter: 80}));
    var authorLine = 'Posted by ' + ((topic.author && topic.author.display_name) || 'Unknown') +
      (topic.posted_at ? '  \u2014  ' + formatDate(topic.posted_at) : '');
    body.push(paragraph(authorLine, {italic: true, size: 20, color: '666666', spacingAfter: 160}));

    var opLines = htmlToLines(topic.message);
    if(!opLines.length) opLines = ['(no text)'];
    opLines.forEach(function(line, i){
      body.push(paragraph(line, {size: 22, spacingAfter: (i === opLines.length - 1) ? 240 : 0}));
    });

    var replyCount = countEntries(view.view);
    body.push(paragraph('Replies (' + replyCount + ')', {bold: true, size: 26, spacingAfter: 120}));

    if(!replyCount){
      body.push(paragraph('No replies yet.', {italic: true, size: 22}));
    } else {
      renderEntries(view.view, 0, body);
    }

    var documentXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      '<w:body>' + body.join('') +
      '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>' +
      '</w:body></w:document>';

    var contentTypesXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
      '</Types>';

    var rootRelsXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
      '</Relationships>';

    var docRelsXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';

    var encoder = new TextEncoder();
    var files = [
      { name: '[Content_Types].xml', data: encoder.encode(contentTypesXml) },
      { name: '_rels/.rels', data: encoder.encode(rootRelsXml) },
      { name: 'word/document.xml', data: encoder.encode(documentXml) },
      { name: 'word/_rels/document.xml.rels', data: encoder.encode(docRelsXml) }
    ];

    var blob = makeZip(files);
    var safeTitle = (topic.title || 'discussion').replace(/[^a-z0-9]+/gi, '-').toLowerCase().slice(0, 60);
    var filename = 'discussion-' + safeTitle + '.docx';

    var url = URL.createObjectURL(blob);
    var a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(function(){ URL.revokeObjectURL(url); }, 1000);

    alert('Exported "' + (topic.title || 'Discussion') + '" (' + replyCount + ' repl' + (replyCount === 1 ? 'y' : 'ies') + ') as ' + filename);
  }).catch(function(err){
    alert('Could not export the discussion: ' + err.message);
  });
}function runGradeReport(){(async()=>{if(document.getElementById("cgr"))return alert("Canvas Grade Report is already open.");const S={courses:[],terms:[],rows:[],term:null};const E=s=>String(s??"").replace(/[&<>"']/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]));const D=s=>s?new Date(s).toLocaleString():"Never/No Activity";const G=(s,g)=>{let a=[];if(s!==null&&s!==undefined&&s!==""){let n=Number(s);if(!isNaN(n))a.push(n.toFixed(2)+"%")}if(g)a.push("("+g+")");return a.join(" ")||"N/A"};async function A(u){let out=[];while(u){let r=await fetch(u,{credentials:"same-origin",headers:{Accept:"application/json"}});if(!r.ok)throw Error("Canvas API "+r.status+": "+u);let d=await r.json();if(!Array.isArray(d))return d;out.push(...d);let l=r.headers.get("Link"),m=l&&l.match(/<([^>]+)>;\s*rel=["']?next["']?/);u=m?m[1]:null}return out}function status(x){document.querySelector("#cgr-status").textContent=x}function prog(n,t){document.querySelector("#cgr-prog").style.width=(t?Math.round(n/t*100):0)+"%"}function est(e){let s=(e.enrollment_state||e.workflow_state||"").toLowerCase();return s==="active"?"Active":s==="inactive"?"Inactive":s==="completed"?"Completed":s==="deleted"?"Withdrawn/Deleted":s||"Unknown"}let o=document.createElement("div");o.id="cgr";o.innerHTML=`<style>#cgr{position:fixed;inset:0;z-index:2147483647;background:#0008;font:14px Arial,sans-serif}#cgr-box{position:absolute;inset:3%;background:#fff;display:flex;flex-direction:column;border-radius:8px;overflow:hidden;box-shadow:0 8px 30px #0008}#cgr-head{background:#2d3b45;color:#fff;padding:14px 18px;display:flex;justify-content:space-between;align-items:center}#cgr-head b{font-size:21px}#cgr-x{font-size:27px;background:none;border:0;color:#fff;cursor:pointer}#cgr-ctl{padding:12px 18px;display:flex;gap:9px;align-items:center;flex-wrap:wrap;border-bottom:1px solid #ccc}#cgr select,#cgr input,#cgr button{padding:8px;font-size:14px}#cgr button{background:#0374b5;color:#fff;border:0;border-radius:4px;cursor:pointer}#cgr button:disabled{opacity:.5}#cgr-status{padding:9px 18px;background:#f4f4f4}.bar{height:7px;background:#ddd}.bar i{display:block;height:100%;width:0;background:#0374b5}#cgr-content{overflow:auto;flex:1;padding:0 16px 18px}.course{margin-top:20px;border:1px solid #cfd6dc;border-radius:6px;overflow:hidden}.course-head{background:#2d3b45;color:#fff;padding:11px 14px}.course-head h3{margin:0 0 5px;font-size:17px;color:#fff}.course-head small{font-size:12px}.tw{overflow:auto}.course table{border-collapse:collapse;width:100%;font-size:13px}.course th{background:#eef1f3;padding:8px;text-align:left;border-bottom:2px solid #999;white-space:nowrap}.course td{padding:7px 8px;border-bottom:1px solid #ddd}.course tr:nth-child(even){background:#fafafa}.inactive{background:#fff7df!important}.withdrawn{background:#f5eeee!important}#cgr-sum{padding:9px 18px;background:#f4f4f4;border-top:1px solid #ccc}</style><div id="cgr-box"><div id="cgr-head"><b>Canvas Grade Report</b><button id="cgr-x">×</button></div><div id="cgr-ctl"><b>Term:</b><select id="cgr-term"><option>Loading...</option></select><button id="cgr-go" disabled>Generate Report</button><input id="cgr-search" placeholder="Search students, status, email..."></div><div id="cgr-status">Loading Canvas courses and terms...</div><div class="bar"><i id="cgr-prog"></i></div><div id="cgr-content"></div><div id="cgr-sum">No report generated.</div></div>`;document.body.appendChild(o);document.querySelector("#cgr-x").onclick=()=>o.remove();function render(){let q=document.querySelector("#cgr-search").value.toLowerCase().trim(),groups=new Map;for(let r of S.rows){if(q&&!r.v.join(" ").toLowerCase().includes(q))continue;if(!groups.has(r.cid))groups.set(r.cid,{name:r.course,rows:[]});groups.get(r.cid).rows.push(r)}let html="";for(let [,g] of groups){g.rows.sort((a,b)=>a.v[0].localeCompare(b.v[0]));let ac=g.rows.filter(r=>r.v[1]==="Active").length,ina=g.rows.filter(r=>r.v[1]==="Inactive").length;html+=`<section class="course"><div class="course-head"><h3>${E(g.name)}</h3><small>${g.rows.length} shown | Active: ${ac} | Inactive: ${ina}</small></div><div class="tw"><table><thead><tr>${["Student","Enrollment Status","Email","Current Grade","Last Activity","Last Submission"].map(x=>`<th>${x}</th>`).join("")}</tr></thead><tbody>${g.rows.map(r=>`<tr class="${r.v[1]==="Inactive"?"inactive":r.v[1]==="Withdrawn/Deleted"?"withdrawn":""}">${r.v.map(v=>`<td>${E(v)}</td>`).join("")}</tr>`).join("")}</tbody></table></div></section>`}document.querySelector("#cgr-content").innerHTML=html||'<p style="padding:20px">No matching records.</p>';let ac=S.rows.filter(r=>r.v[1]==="Active").length,ina=S.rows.filter(r=>r.v[1]==="Inactive").length;document.querySelector("#cgr-sum").textContent=`All courses: ${S.rows.length} records | Active: ${ac} | Inactive: ${ina}`}document.querySelector("#cgr-ctl").insertAdjacentHTML("beforeend",'<button id="cgr-export" disabled>Export to .xls</button>');
function sanitizeSheetName(name,used){let clean=String(name).replace(/[:\\\/\?\*\[\]]/g," ").trim();if(clean.length>31)clean=clean.slice(0,31);if(!clean)clean="Sheet";let candidate=clean,n=2;while(used[candidate]){let suffix=" ("+n+")";candidate=clean.slice(0,31-suffix.length)+suffix;n++}used[candidate]=true;return candidate}
function exportXls(){
  if(!S.rows.length)return alert("Generate a report first.");
  let usedNames={};
  let xml='<?xml version="1.0"?><?mso-application progid="Excel.Sheet"?><Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet" xmlns:html="http://www.w3.org/TR/REC-html40"><Styles><Style ss:ID="Header"><Font ss:Bold="1"/></Style></Styles>';
  let groups=new Map();
  for(let r of S.rows){if(!groups.has(r.cid))groups.set(r.cid,{name:r.course,rows:[]});groups.get(r.cid).rows.push(r)}
  let summaryName=sanitizeSheetName("Summary",usedNames);
  xml+=`<Worksheet ss:Name="${summaryName}"><Table><Row><Cell ss:StyleID="Header"><Data ss:Type="String">Course</Data></Cell><Cell ss:StyleID="Header"><Data ss:Type="String">Active</Data></Cell><Cell ss:StyleID="Header"><Data ss:Type="String">Inactive</Data></Cell><Cell ss:StyleID="Header"><Data ss:Type="String">Total</Data></Cell></Row>`;
  for(let [,g] of groups){
    let ac=g.rows.filter(r=>r.v[1]==="Active").length,ina=g.rows.filter(r=>r.v[1]==="Inactive").length;
    xml+=`<Row><Cell><Data ss:Type="String">${E(g.name)}</Data></Cell><Cell><Data ss:Type="Number">${ac}</Data></Cell><Cell><Data ss:Type="Number">${ina}</Data></Cell><Cell><Data ss:Type="Number">${g.rows.length}</Data></Cell></Row>`;
  }
  xml+="</Table></Worksheet>";
  let headers=["Student","Enrollment Status","Email","Current Grade","Last Activity","Last Submission"];
  for(let [,g] of groups){
    let sheetName=sanitizeSheetName(g.name,usedNames);
    xml+=`<Worksheet ss:Name="${sheetName}"><Table><Row>`+headers.map(h=>`<Cell ss:StyleID="Header"><Data ss:Type="String">${E(h)}</Data></Cell>`).join("")+"</Row>";
    let sorted=g.rows.slice().sort((a,b)=>a.v[0].localeCompare(b.v[0]));
    for(let r of sorted){
      xml+="<Row>"+r.v.map(v=>`<Cell><Data ss:Type="String">${E(v)}</Data></Cell>`).join("")+"</Row>";
    }
    xml+="</Table></Worksheet>";
  }
  xml+="</Workbook>";
  let blob=new Blob([xml],{type:"application/vnd.ms-excel"});
  let url=URL.createObjectURL(blob);
  let a=document.createElement("a");
  a.href=url;
  a.download="canvas-grade-report-"+(S.term?S.term.name.replace(/[^a-z0-9]+/gi,"-").toLowerCase():"report")+".xls";
  document.body.appendChild(a);a.click();document.body.removeChild(a);
  setTimeout(()=>URL.revokeObjectURL(url),1000);
}
document.querySelector("#cgr-export").onclick=exportXls;
document.querySelector("#cgr-search").oninput=render;try{S.courses=await A("/api/v1/courses?enrollment_type=teacher&state[]=available&include[]=term&per_page=100");let m=new Map;for(let c of S.courses){if(!c.term||!c.term.id)continue;let t=m.get(c.term.id)||{id:c.term.id,name:c.term.name||("Term "+c.term.id),start:c.term.start_at,end:c.term.end_at,courses:[]};t.courses.push(c);m.set(c.term.id,t)}let now=Date.now();S.terms=[...m.values()].filter(t=>!t.end||isNaN(new Date(t.end))||new Date(t.end).getTime()>=now).sort((a,b)=>(new Date(b.start||0))-(new Date(a.start||0)));if(!S.terms.length)S.terms=[...m.values()];let sel=document.querySelector("#cgr-term");sel.innerHTML='<option value="">Select a term...</option>'+S.terms.map(t=>`<option value="${t.id}">${E(t.name)} (${t.courses.length} course${t.courses.length===1?"":"s"})</option>`).join("");document.querySelector("#cgr-go").disabled=false;status("Select a term and click Generate Report.")}catch(e){console.error(e);status("Unable to load courses: "+e.message);return}document.querySelector("#cgr-go").onclick=async()=>{let id=document.querySelector("#cgr-term").value;if(!id)return alert("Select a term.");S.term=S.terms.find(t=>String(t.id)===id);S.rows=[];render();let btn=document.querySelector("#cgr-go");btn.disabled=true;let errs=0,n=0;for(let c of S.term.courses){status("Processing "+(c.name||c.course_code||c.id)+"...");try{let ep=new URLSearchParams;ep.append("type[]","StudentEnrollment");["active","inactive"].forEach(x=>ep.append("state[]",x));ep.append("include[]","email");ep.append("per_page","100");let sp=new URLSearchParams;sp.append("student_ids[]","all");sp.append("per_page","100");let [en,sub]=await Promise.all([A(`/api/v1/courses/${c.id}/enrollments?${ep}`),A(`/api/v1/courses/${c.id}/students/submissions?${sp}`)]);let lm=new Map;for(let s of sub){if(!s.user_id||!s.submitted_at)continue;let z=lm.get(String(s.user_id));if(!z||new Date(s.submitted_at)>new Date(z))lm.set(String(s.user_id),s.submitted_at)}for(let e of en){if(!e.user||!["Active","Inactive"].includes(est(e)))continue;let u=e.user,g=e.grades||{},ls=lm.get(String(u.id))||"";S.rows.push({cid:c.id,course:c.name||c.course_code||("Course "+c.id),v:[u.sortable_name||u.name||("User "+u.id),est(e),u.email||u.login_id||"N/A",G(g.current_score,g.current_grade),D(e.last_activity_at),D(ls)]})}}catch(e){errs++;console.error("Canvas Grade Report course error",c,e)}prog(++n,S.term.courses.length);render()}status(`Report complete. ${S.rows.length} student-course records loaded${errs?`; ${errs} course(s) had API errors. See Console.`:"."}`);btn.disabled=false;document.querySelector("#cgr-export").disabled=!S.rows.length}})();}function runPreNoShowReport(){(async function(){'use strict';const V='1.5',LK='preNoShowLog_v1',Z=2147483000,H=window.location.hostname,O=window.location.origin,ATTENDANCE_NAME_MATCH='attendance';async function cfa(u){const r=[];let n=u.indexOf('per_page=')!==-1?u:u+(u.indexOf('?')!==-1?'&':'?')+'per_page=100';while(n){const res=await fetch(n,{credentials:'include',headers:{'Accept':'application/json'}});if(!res.ok){if(res.status===401)throw new Error('SESSION_EXPIRED');throw new Error('HTTP '+res.status);}const d=await res.json();if(Array.isArray(d))r.push(...d);else return d;const lh=res.headers.get('Link')||'';let nextUrl=null;const parts=lh.split(',');for(let i=0;i<parts.length;i++){if(parts[i].indexOf('rel="next"')!==-1){const sIdx=parts[i].indexOf('<'),eIdx=parts[i].indexOf('>');if(sIdx!==-1&&eIdx!==-1)nextUrl=parts[i].substring(sIdx+1,eIdx);}}n=nextUrl;}return r;}async function cfo(u){const r=await fetch(u,{credentials:'include',headers:{'Accept':'application/json'}});if(!r.ok){if(r.status===401)throw new Error('SESSION_EXPIRED');throw new Error('HTTP '+r.status);}return r.json();}function fd(seconds){if(!seconds||seconds<0)return'0h 0m';const h=Math.floor(seconds/3600);const m=Math.floor((seconds-(h*3600))/60);return h+'h '+m+'m';}function fla(i){if(!i)return{display:'Never',tooltip:'No recorded access'};const t=new Date(i),n=new Date(),d=Math.floor((n-t)/86400000);let x;if(d===0)x='Today';else if(d===1)x='Yesterday';else if(d<30)x=d+' days ago';else x=t.toLocaleDateString();return{display:x,tooltip:t.toLocaleString()};}function esc(s){if(s==null)return'';return String(s).split('&').join('&amp;').split('<').join('&lt;').split('>').join('&gt;').split('"').join('&quot;').split("'").join('&#39;');}function ll(){try{return JSON.parse(localStorage.getItem(LK)||'{}');}catch(e){return{};}}function sl(l){localStorage.setItem(LK,JSON.stringify(l));}function lk(c,u){return'course_'+c+'_user_'+u;}function rm(){document.querySelectorAll('.pns-modal,.pns-overlay,.pns-progress,.pns-backdrop,.pns-toast,.pns-compose-backdrop').forEach(e=>e.remove());}function is(){if(document.getElementById('pns-styles'))return;const s=document.createElement('style');s.id='pns-styles';const rawCss='.pns-backdrop{position:fixed;top:0;left:0;width:100vw;height:100vh;background:rgba(0,0,0,0.65);z-index:'+Z+';display:flex;align-items:center;justify-content:center;font-family:sans-serif;}.pns-modal{background:#fff;padding:24px 28px;border-radius:8px;box-shadow:0 10px 40px rgba(0,0,0,0.45);min-width:400px;max-width:600px;}.pns-modal h2{margin:0 0 16px 0;color:#2d3b45;font-size:20px;}.pns-modal label{display:block;margin:12px 0 6px;font-weight:600;color:#2d3b45;}.pns-modal select{width:100__PCT__;padding:8px;font-size:14px;border:1px solid #c7cdd1;border-radius:4px;background:#fff;}.pns-btn{padding:8px 18px;font-size:14px;border-radius:4px;cursor:pointer;border:1px solid transparent;margin-left:8px;font-weight:600;}.pns-btn-primary{background:#0374B5;color:#fff;border-color:#0374B5;}.pns-btn-primary:hover{background:#025a8c;}.pns-btn-secondary{background:#fff;color:#2d3b45;border-color:#c7cdd1;}.pns-btn-danger{background:#c72d2d;color:#fff;border-color:#c72d2d;}.pns-btn-row{text-align:right;margin-top:20px;}.pns-progress{background:#fff;padding:24px 28px;border-radius:8px;min-width:420px;text-align:center;}.pns-progress .pns-spinner{width:36px;height:36px;margin:0 auto 14px;border:4px solid #e0e0e0;border-top-color:#0374B5;border-radius:999px;animation:pns-spin 1s linear infinite;}@keyframes pns-spin{to{transform:rotate(360deg);}}.pns-progress .pns-status{font-size:14px;color:#2d3b45;margin-bottom:10px;min-height:20px;}.pns-progress .pns-bar-wrap{background:#e0e0e0;height:10px;border-radius:5px;overflow:hidden;}.pns-progress .pns-bar{background:#0374B5;height:10px;width:0;transition:width 0.3s;}.pns-overlay{position:fixed;top:0;left:0;width:100vw;height:100vh;background:#fff;z-index:'+Z+';overflow:auto;font-family:sans-serif;color:#2d3b45;}.pns-report{max-width:1200px;margin:0 auto;padding:24px;}.pns-header{border-bottom:2px solid #2d3b45;padding-bottom:16px;margin-bottom:24px;display:flex;justify-content:space-between;align-items:flex-start;}.pns-header h1{margin:0 0 6px 0;font-size:24px;}.pns-header .pns-meta{font-size:13px;color:#556;}.pns-header-actions{display:flex;gap:8px;}.pns-course-block{margin-bottom:36px;page-break-inside:avoid;}.pns-course-block h3{background:#f5f5f5;padding:10px 14px;border-left:4px solid #0374B5;margin:0 0 8px 0;font-size:16px;}.pns-course-block .pns-course-meta{font-size:12px;color:#666;margin-bottom:8px;padding-left:14px;}table.pns-table{width:100__PCT__;border-collapse:collapse;font-size:13px;}table.pns-table th{background:#2d3b45;color:#fff;padding:8px 6px;text-align:left;font-weight:600;}table.pns-table td{padding:6px;border-bottom:1px solid #e0e0e0;vertical-align:middle;}table.pns-table td.pns-num{font-family:monospace;text-align:right;}table.pns-table tr:hover{background:#f9f9f9;}.pns-action-link{color:#c72d2d;font-weight:600;text-decoration:none;padding:4px 8px;border:1px solid #c72d2d;border-radius:3px;font-size:12px;white-space:nowrap;display:inline-block;cursor:pointer;}.pns-action-link:hover{background:#c72d2d;color:#fff;}.pns-action-messaged{color:#8a6d3b;font-weight:600;padding:4px 8px;border:1px solid #8a6d3b;border-radius:3px;font-size:12px;background:#fcf8e3;cursor:pointer;white-space:nowrap;display:inline-block;text-decoration:none;}.pns-action-none{color:#999;}.pns-footer{margin-top:36px;padding-top:16px;border-top:1px solid #ccc;font-size:11px;color:#666;line-height:1.5;}.pns-error-note{background:#fcf2f2;border-left:4px solid #c72d2d;padding:8px 12px;margin:10px 0;font-size:12px;color:#7a1f1f;}.pns-toast{position:fixed;bottom:24px;left:50__PCT__;transform:translateX(-50__PCT__);background:#2d3b45;color:#fff;padding:14px 22px;border-radius:8px;box-shadow:0 4px 20px rgba(0,0,0,0.35);z-index:'+(Z+20)+';font-family:sans-serif;font-size:14px;max-width:500px;text-align:center;line-height:1.5;}.pns-toast strong{color:#a6e3a1;}.pns-toast .pns-toast-close{margin-left:14px;color:#fff;cursor:pointer;font-weight:bold;background:transparent;border:1px solid #fff;border-radius:3px;padding:2px 8px;font-size:12px;}.pns-compose-backdrop{position:fixed;top:0;left:0;width:100vw;height:100vh;background:rgba(0,0,0,0.6);z-index:'+(Z+10)+';display:flex;align-items:center;justify-content:center;font-family:sans-serif;}.pns-compose-modal{background:#fff;padding:24px;border-radius:8px;box-shadow:0 10px 40px rgba(0,0,0,0.4);width:650px;max-width:90__PCT__;display:flex;flex-direction:column;gap:12px;}.pns-compose-modal h3{margin:0 0 4px 0;color:#2d3b45;font-size:18px;}.pns-compose-modal label{font-weight:600;font-size:13px;color:#444;display:block;margin-bottom:4px;}.pns-compose-field{width:100__PCT__;padding:8px;font-size:14px;border:1px solid #c7cdd1;border-radius:4px;font-family:sans-serif;}.pns-compose-body{height:250px;resize:vertical;}@media print{.pns-header-actions,.pns-action-link,.pns-action-messaged{display:none !important;}.pns-overlay{position:static;overflow:visible;}.pns-course-block{page-break-inside:avoid;}}';s.textContent=rawCss.split('__PCT__').join(String.fromCharCode(37));document.head.appendChild(s);}let pc=false;function sp(){pc=false;rm();const b=document.createElement('div');b.className='pns-backdrop';b.innerHTML='<div class="pns-progress"><div class="pns-spinner"></div><div class="pns-status" id="pns-status">Initializing…</div><div class="pns-bar-wrap"><div class="pns-bar" id="pns-bar"></div></div><div class="pns-btn-row" style="text-align:center;"><button class="pns-btn pns-btn-secondary" id="pns-cancel">Cancel</button></div></div>';document.body.appendChild(b);document.getElementById('pns-cancel').onclick=()=>{pc=true;rm();};}function up(p,m){const b=document.getElementById('pns-bar'),s=document.getElementById('pns-status');if(b)b.style.width=p+String.fromCharCode(37);if(s)s.textContent=m;}function showToast(msg,duration){const existing=document.querySelector('.pns-toast');if(existing)existing.remove();const t=document.createElement('div');t.className='pns-toast';t.innerHTML=msg+' <button class="pns-toast-close">Dismiss</button>';document.body.appendChild(t);t.querySelector('.pns-toast-close').onclick=()=>t.remove();if(duration)setTimeout(()=>{if(t.parentNode)t.remove();},duration);}function messageHref(courseId,uid,name){return O+'/conversations?context_id=course_'+courseId+'&user_id='+encodeURIComponent(uid)+'&user_name='+encodeURIComponent(name);}function legacyCopy(ta){try{ta.focus();ta.select();return document.execCommand('copy');}catch(e){return false;}}function setNativeValue(win,el,val){const proto=el.tagName==='TEXTAREA'?win.HTMLTextAreaElement.prototype:win.HTMLInputElement.prototype;const setter=Object.getOwnPropertyDescriptor(proto,'value').set;setter.call(el,val);el.dispatchEvent(new win.Event('input',{bubbles:true}));el.dispatchEvent(new win.Event('change',{bubbles:true}));}function tryAutoFillMessage(win,subjectText,bodyText,onDone){let attempts=0;const maxAttempts=40;const timer=setInterval(()=>{attempts++;if(win.closed){clearInterval(timer);onDone(false,false);return;}let textarea=null,subjectInput=null;try{const doc=win.document;textarea=doc.querySelector('textarea[data-testid="message-body"]')||doc.querySelector('textarea[name="body"]');if(!textarea){const areas=Array.prototype.slice.call(doc.querySelectorAll('textarea'));const visible=areas.filter(t=>t.offsetWidth>0&&t.offsetHeight>0);if(visible.length){visible.sort((a,b)=>(b.offsetWidth*b.offsetHeight)-(a.offsetWidth*a.offsetHeight));textarea=visible[0];}}subjectInput=doc.querySelector('input[data-testid="subject-input"]')||doc.querySelector('input[name="subject"]')||doc.querySelector('input[placeholder="Subject"]');}catch(e){textarea=null;}if(textarea){clearInterval(timer);setNativeValue(win,textarea,bodyText);if(subjectInput)setNativeValue(win,subjectInput,subjectText);onDone(true,!!subjectInput);return;}if(attempts>=maxAttempts){clearInterval(timer);onDone(false,false);}},200);}async function stp(cs,iN){const tm=new Map();cs.forEach(c=>{if(c.term&&c.term.id)tm.set(c.term.id,{id:c.term.id,name:c.term.name||'Unnamed Term',start_at:c.term.start_at||'0'});});const ts=Array.from(tm.values()).sort((a,b)=>(b.start_at||'').localeCompare(a.start_at||''));if(!ts.length){alert('No terms found.');return null;}return new Promise(res=>{rm();const b=document.createElement('div');b.className='pns-backdrop';b.innerHTML='<div class="pns-modal"><h2>Pre-No-Show Report</h2><p style="margin:0 0 12px;font-size:13px;color:#556;">Instructor: <strong>'+esc(iN)+'</strong><br>Select a term to generate the participation report.</p><label for="pns-term-sel">Term:</label><select id="pns-term-sel">'+ts.map(t=>'<option value="'+esc(t.id)+'">'+esc(t.name)+'</option>').join('')+'</select><div class="pns-btn-row"><button class="pns-btn pns-btn-secondary" id="pns-cancel-term">Cancel</button><button class="pns-btn pns-btn-primary" id="pns-go">Generate Report</button></div></div>';document.body.appendChild(b);document.getElementById('pns-cancel-term').onclick=()=>{rm();res(null);};document.getElementById('pns-go').onclick=()=>{const v=document.getElementById('pns-term-sel').value,ch=ts.find(t=>String(t.id)===String(v));rm();res(ch);};});}function classifySubmission(s){const a=s.assignment;const isQuiz=s.submission_type==='online_quiz'||(a&&a.is_quiz_assignment===true);const isLTI=!isQuiz&&a&&Array.isArray(a.submission_types)&&a.submission_types.indexOf('external_tool')!==-1;return{isQuiz:isQuiz,isLTI:isLTI};}function collectActivityScores(sb,matchFn,now){const assignments=new Map();const scores=new Map();sb.forEach(s=>{const a=s.assignment;if(!a||!matchFn(a))return;const due=a.due_at?new Date(a.due_at).getTime():null;if(due!==null&&due>now)return;if(!(a.points_possible>0))return;if(!assignments.has(a.id))assignments.set(a.id,a.points_possible);if(s.score!=null){const pct=Math.max(0,Math.min(100,(s.score/a.points_possible)*100));scores.set(a.id+'_'+s.user_id,pct);}});return{assignments:assignments,scores:scores};}function avgPctFor(userId,activity){if(activity.assignments.size===0)return null;let sum=0;activity.assignments.forEach((pointsPossible,assignmentId)=>{const key=assignmentId+'_'+userId;sum+=activity.scores.has(key)?activity.scores.get(key):0;});return sum/activity.assignments.size;}async function gdc(cid){const cs={};const tp=await cfa(O+'/api/v1/courses/'+cid+'/discussion_topics?per_page=100');const B=5;for(let i=0;i<tp.length;i+=B){const bt=tp.slice(i,i+B);await Promise.all(bt.map(async t=>{try{const es=await cfa(O+'/api/v1/courses/'+cid+'/discussion_topics/'+t.id+'/entries?per_page=100');const deep=[];es.forEach(e=>{cs[e.user_id]=(cs[e.user_id]||0)+1;if(e.has_more_replies){deep.push(e.id);}else if(e.recent_replies){e.recent_replies.forEach(r=>{cs[r.user_id]=(cs[r.user_id]||0)+1;});}});for(let j=0;j<deep.length;j+=B){const db=deep.slice(j,j+B);await Promise.all(db.map(async eid=>{try{const reps=await cfa(O+'/api/v1/courses/'+cid+'/discussion_topics/'+t.id+'/entries/'+eid+'/replies?per_page=100');reps.forEach(r=>{cs[r.user_id]=(cs[r.user_id]||0)+1;});}catch(er){}}));}}catch(er){}}));}return cs;}async function gcd(c){const cid=c.id,er=[];const[en,sb,de]=await Promise.all([cfa(O+'/api/v1/courses/'+cid+'/enrollments?type[]=StudentEnrollment&state[]=active&include[]=total_scores').catch(e=>{er.push('enrollments');return[];}),cfa(O+'/api/v1/courses/'+cid+'/students/submissions?student_ids[]=all&include[]=assignment&per_page=100').catch(e=>{er.push('submissions');return[];}),gdc(cid).catch(e=>{er.push('discussions');return{};})]);const sm=new Map();en.forEach(e=>{const u=e.user_id,us=e.user||{};sm.set(u,{user_id:u,name:us.sortable_name||us.name||'User '+u,short_name:us.short_name||us.name||'',last_activity_at:e.last_activity_at,total_activity_time:e.total_activity_time||0,current_score:(e.grades&&e.grades.current_score!==undefined)?e.grades.current_score:null,assignments_submitted:0,quizzes_submitted:0,lti_submitted:0,discussion_posts:de[u]||0,lti_avg_pct:null,attendance_pct:null});});sb.forEach(s=>{const st=sm.get(s.user_id);if(!st)return;if(s.workflow_state==='unsubmitted'||!s.submitted_at)return;const cls=classifySubmission(s);if(cls.isQuiz)st.quizzes_submitted++;else if(cls.isLTI)st.lti_submitted++;else st.assignments_submitted++;});const now=Date.now();const ltiActivity=collectActivityScores(sb,a=>Array.isArray(a.submission_types)&&a.submission_types.indexOf('external_tool')!==-1&&a.is_quiz_assignment!==true,now);const attendanceActivity=collectActivityScores(sb,a=>typeof a.name==='string'&&a.name.toLowerCase().indexOf(ATTENDANCE_NAME_MATCH)!==-1,now);sm.forEach(st=>{st.lti_avg_pct=avgPctFor(st.user_id,ltiActivity);st.attendance_pct=avgPctFor(st.user_id,attendanceActivity);});return{course:c,students:Array.from(sm.values()).sort((a,b)=>a.name.localeCompare(b.name)),errors:er};}function buildSubject(c){return'Urgent: Course Participation Required — '+c.name;}function buildBody(s,c,iN){const fn=(s.short_name||s.name.split(',').pop()||'').trim().split(' ')[0];const nl=String.fromCharCode(10);return'Hi '+fn+','+nl+nl+'I\'m reaching out because I\'ve noticed that you have not yet participated in '+c.name+' this term. Specifically, my records show that you have not:'+nl+nl+'  • Posted to any discussion boards'+nl+'  • Submitted any assignments'+nl+'  • Completed any quizzes'+nl+'  • Completed any interactive/LTI-based course activities'+nl+nl+'Because of this lack of participation, you are currently at risk of being reported to the Financial Aid Office as a no-show for this course. A no-show designation can affect your financial aid eligibility and enrollment status.'+nl+nl+'I want to help you succeed. Please take one of the following steps within the next 48 hours:'+nl+nl+'  1. Log into Canvas and submit any outstanding work, OR'+nl+'  2. Reply to this message to let me know your situation, OR'+nl+'  3. Contact me during office hours to discuss.'+nl+nl+'If you no longer plan to take this course, please contact the Registrar\'s Office immediately to formally withdraw so this does not impact your academic record.'+nl+nl+'I look forward to hearing from you.'+nl+nl+'Best regards,'+nl+iN;}function ins(s){return s.discussion_posts===0&&s.quizzes_submitted===0&&s.assignments_submitted===0&&s.lti_submitted===0&&(s.current_score===0||s.current_score===null);}function csvEsc(v){if(v==null)return'';const s=String(v);if(s.indexOf('"')!==-1||s.indexOf(',')!==-1||s.indexOf(String.fromCharCode(10))!==-1||s.indexOf(String.fromCharCode(13))!==-1){return'"'+s.split('"').join('""')+'"';}return s;}function buildCSVRows(rs){const lg=ll();const rows=[['Course','Course Code','Student','Student ID','Last Access','Discussion Posts','Activity','Assignments Submitted','Quizzes Submitted','LTI Activities %','Attendance %','Current Score','Flagged','Draft Opened Date']];rs.forEach(r=>{const c=r.course;r.students.forEach(s=>{const la=fla(s.last_activity_at);const fl=ins(s);const gd=(s.current_score==null)?'N/A':s.current_score+String.fromCharCode(37);const ltiPct=(s.lti_avg_pct==null)?'N/A':Math.round(s.lti_avg_pct)+String.fromCharCode(37);const attPct=(s.attendance_pct==null)?'N/A':Math.round(s.attendance_pct)+String.fromCharCode(37);const k=lk(c.id,s.user_id);const pm=lg[k];rows.push([c.name,c.course_code||'',s.name,s.user_id,la.display,s.discussion_posts,fd(s.total_activity_time),s.assignments_submitted,s.quizzes_submitted,ltiPct,attPct,gd,fl?'Yes':'No',pm?new Date(pm.messaged_at).toLocaleDateString():'']);});});return rows;}function downloadCSV(tN,rs){const rows=buildCSVRows(rs);const csv=rows.map(row=>row.map(csvEsc).join(',')).join(String.fromCharCode(13)+String.fromCharCode(10));const cleanTerm=tN.split(' ').join('');const fname='PreNoShow_'+cleanTerm+'_'+new Date().toISOString().slice(0,10)+'.csv';const blob=new Blob([csv],{type:'text/csv;charset=utf-8;'});const url=URL.createObjectURL(blob);const a=document.createElement('a');a.href=url;a.download=fname;document.body.appendChild(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(url),1000);}function openComposeModal(student,course,instructorName,linkEl){return new Promise(resolve=>{const subj=buildSubject(course);const body=buildBody(student,course,instructorName);const bdrop=document.createElement('div');bdrop.className='pns-compose-backdrop';bdrop.innerHTML='<div class="pns-compose-modal"><h3>✉️ Compose Financial Aid Warning</h3><div><label>To (Student):</label><input type="text" class="pns-compose-field" value="'+esc(student.name)+'" readonly style="background:#f5f5f5;font-weight:600;"></div><div><label>Course Context:</label><input type="text" class="pns-compose-field" value="'+esc(course.name)+'" readonly style="background:#f5f5f5;"></div><div><label>Subject:</label><input type="text" id="pns-comp-subj" class="pns-compose-field" value="'+esc(subj)+'"></div><div><label>Message Body:</label><textarea id="pns-comp-body" class="pns-compose-field pns-compose-body">'+esc(body)+'</textarea></div><div class="pns-btn-row" id="pns-comp-actions"><button class="pns-btn pns-btn-secondary" id="pns-comp-cancel">Cancel</button><button class="pns-btn pns-btn-primary" id="pns-comp-send">Open in Canvas Message</button></div></div>';document.body.appendChild(bdrop);document.getElementById('pns-comp-cancel').onclick=()=>{bdrop.remove();resolve(false);};document.getElementById('pns-comp-send').onclick=()=>{const actionsRow=document.getElementById('pns-comp-actions');const subjVal=document.getElementById('pns-comp-subj').value;const bodyVal=document.getElementById('pns-comp-body').value;function logAndUpdate(){const cL=ll();cL[lk(course.id,student.user_id)]={messaged_at:new Date().toISOString(),term:linkEl.dataset.term,student_name:student.name};sl(cL);linkEl.className='pns-action-messaged';linkEl.textContent='✎ Drafted '+new Date().toLocaleDateString();}function clipboardFallback(){function done(ok){bdrop.remove();showToast(ok?'<strong>Canvas message opened</strong> and the draft was copied to your clipboard for <strong>'+esc(student.name)+'</strong> — paste it into the message body, then review and click Send yourself.':'<strong>Canvas message opened</strong> for <strong>'+esc(student.name)+'</strong>. Could not copy automatically — copy the text from this editor and paste it in yourself.',6000);logAndUpdate();resolve(true);}if(navigator.clipboard&&navigator.clipboard.writeText){navigator.clipboard.writeText(bodyVal).then(()=>done(true),()=>done(legacyCopy(document.getElementById('pns-comp-body'))));}else{done(legacyCopy(document.getElementById('pns-comp-body')));}}const win=window.open(messageHref(course.id,student.user_id,student.name),'_blank');if(win){actionsRow.innerHTML='<span style="font-size:14px;color:#555;font-weight:600;">Opening Canvas message…</span>';tryAutoFillMessage(win,subjVal,bodyVal,(filled)=>{if(filled){bdrop.remove();showToast('<strong>Canvas message opened and filled in</strong> for <strong>'+esc(student.name)+'</strong> — review it in the tab that just opened, then click Send yourself.',6000);logAndUpdate();resolve(true);}else{clipboardFallback();}});}else{clipboardFallback();}};});}function rr(tN,iN,rs){rm();const lg=ll(),n=new Date(),ds=n.toLocaleString();const cleanTerm=tN.split(' ').join('');const fh='PreNoShow_'+cleanTerm+'_'+n.toISOString().slice(0,10)+'.pdf';let tS=0,tF=0;rs.forEach(r=>{tS+=r.students.length;tF+=r.students.filter(ins).length;});const ov=document.createElement('div');ov.className='pns-overlay';const studentDataMap={};let h='<div class="pns-report"><div class="pns-header"><div><h1>Pre-No-Show Report — '+esc(tN)+'</h1><div class="pns-meta">Generated: '+esc(ds)+'<br>Instructor: <strong>'+esc(iN)+'</strong><br>'+rs.length+' course(s) • '+tS+' student(s) • <strong style="color:#c72d2d;">'+tF+' flagged</strong></div></div><div class="pns-header-actions"><button class="pns-btn pns-btn-secondary" id="pns-export-csv">Export CSV</button><button class="pns-btn pns-btn-secondary" id="pns-clear-log">Clear Message Log</button><button class="pns-btn pns-btn-primary" id="pns-print">Print / Save PDF</button><button class="pns-btn pns-btn-danger" id="pns-close">Close</button></div></div><div class="pns-error-note" style="display:'+(rs.some(r=>r.errors.length)?'block':'none')+';">Some course data may be incomplete.</div><div style="font-size:11px;color:#666;margin-bottom:20px;">Suggested filename: <code>'+esc(fh)+'</code><br><strong>Note:</strong> Clicking "Draft Warning" opens an editable message, then hands off to Canvas’s own Inbox with the subject and body filled in (or copied to your clipboard if auto-fill isn’t available) — review it there and click Send yourself. This tool never sends anything automatically.<br><strong>Attendance %</strong> is matched by assignment name containing "'+esc(ATTENDANCE_NAME_MATCH)+'" — unconfirmed against your actual AttendancePlus setup; verify before relying on this column.</div>';rs.forEach(r=>{const c=r.course,fc=r.students.filter(ins).length;h+='<div class="pns-course-block"><h3>📚 '+esc(c.name)+(c.course_code?' ('+esc(c.course_code)+')':'')+'</h3><div class="pns-course-meta">'+r.students.length+' student(s) • '+fc+' flagged'+(r.errors.length?' • <span style="color:#c72d2d;">⚠️ '+r.errors.join(', ')+'</span>':'')+'</div><table class="pns-table"><thead><tr><th>Student</th><th>Last Access</th><th style="text-align:right;">Posts</th><th style="text-align:right;">Activity</th><th style="text-align:right;">Asgn</th><th style="text-align:right;">Qz</th><th style="text-align:right;">LTI %</th><th style="text-align:right;">Attend %</th><th style="text-align:right;">Grade</th><th>Action</th></tr></thead><tbody>';r.students.forEach(s=>{const la=fla(s.last_activity_at),fl=ins(s),gd=(s.current_score==null)?'N/A':s.current_score+String.fromCharCode(37);const ltiPct=(s.lti_avg_pct==null)?'N/A':Math.round(s.lti_avg_pct)+String.fromCharCode(37);const attPct=(s.attendance_pct==null)?'N/A':Math.round(s.attendance_pct)+String.fromCharCode(37);let ah='<span class="pns-action-none">—</span>';if(fl){const k=lk(c.id,s.user_id),pm=lg[k];const dataKey=c.id+'_'+s.user_id;studentDataMap[dataKey]={student:s,course:c};if(pm){const md=new Date(pm.messaged_at).toLocaleDateString();ah='<a href="#" class="pns-action-messaged" data-key="'+esc(dataKey)+'" data-cid="'+esc(c.id)+'" data-uid="'+esc(s.user_id)+'" data-sname="'+esc(s.name)+'" data-term="'+esc(tN)+'">⚠️ Messaged '+esc(md)+'</a>';}else{ah='<a href="#" class="pns-action-link" data-key="'+esc(dataKey)+'" data-cid="'+esc(c.id)+'" data-uid="'+esc(s.user_id)+'" data-sname="'+esc(s.name)+'" data-term="'+esc(tN)+'">Draft Warning</a>';}}h+='<tr'+(fl?' style="background:#fef5f5;"':'')+'><td>'+esc(s.name)+'</td><td title="'+esc(la.tooltip)+'">'+esc(la.display)+'</td><td class="pns-num">'+s.discussion_posts+'</td><td class="pns-num">'+esc(fd(s.total_activity_time))+'</td><td class="pns-num">'+s.assignments_submitted+'</td><td class="pns-num">'+s.quizzes_submitted+'</td><td class="pns-num">'+esc(ltiPct)+'</td><td class="pns-num">'+esc(attPct)+'</td><td class="pns-num">'+esc(gd)+'</td><td>'+ah+'</td></tr>';});h+='</tbody></table></div>';});h+='<div class="pns-footer">Pre-No-Show Bookmarklet v'+V+' • '+esc(H)+' • '+esc(ds)+'<br><em>Activity Time is an estimate. Consult official no-show procedures.</em></div></div>';ov.innerHTML=h;document.body.appendChild(ov);document.getElementById('pns-close').onclick=()=>rm();document.getElementById('pns-print').onclick=()=>window.print();document.getElementById('pns-export-csv').onclick=()=>{try{downloadCSV(tN,rs);}catch(err){alert('CSV export failed: '+err.message);}};document.getElementById('pns-clear-log').onclick=()=>{if(confirm('Clear message log?')){localStorage.removeItem(LK);alert('Cleared. Close and re-run to refresh.');}};ov.querySelectorAll('.pns-action-link,.pns-action-messaged').forEach(lnk=>{lnk.addEventListener('click',async function(ev){ev.preventDefault();const dataKey=this.dataset.key;const data=studentDataMap[dataKey];if(!data){alert('Error: student data not found.');return;}if(this.classList.contains('pns-action-messaged')&&!confirm('Already drafted a warning for '+this.dataset.sname+'. Draft again?')){return;}await openComposeModal(data.student,data.course,iN,this);});});}try{is();sp();up(2,'Loading profile…');const me=await cfo(O+'/api/v1/users/self');const iN=me.name||'Instructor';up(5,'Loading courses…');const ac=await cfa(O+'/api/v1/users/self/courses?enrollment_type=teacher&state[]=available&include[]=term');if(pc)return;if(!ac.length){rm();alert('No courses found.');return;}rm();const st=await stp(ac,iN);if(!st)return;const tc=ac.filter(c=>c.term&&String(c.term.id)===String(st.id)&&c.workflow_state==='available');if(!tc.length){alert('No courses in selected term.');return;}sp();const rs=[];for(let i=0;i<tc.length;i++){if(pc)return;const currentPct=Math.round((i/tc.length)*90)+10;up(currentPct,'Course '+(i+1)+'/'+tc.length+': '+tc[i].name);try{rs.push(await gcd(tc[i]));}catch(er){if(er.message==='SESSION_EXPIRED'){rm();alert('Session expired.');return;}rs.push({course:tc[i],students:[],errors:['error']});}}if(pc)return;up(100,'Rendering…');await new Promise(r=>setTimeout(r,200));rr(st.name,iN,rs);}catch(er){rm();console.error(er);alert(er.message==='SESSION_EXPIRED'?'Session expired.':'Error: '+er.message);}})();}


  function removeOverlay(id){
    var old = document.getElementById(id);
    if(old) old.remove();
  }
  function esc(s){
    return String(s == null ? '' : s).replace(/[&<>]/g, function(c){ return {'&':'&amp;','<':'&lt;','>':'&gt;'}[c]; });
  }

  var tools = [
    { label: 'Syllabus Auto-Builder', desc: 'Weekly schedule, assignment list, and grading breakdown for this course.', course: true, fn: runSyllabusBuilder },
    { label: 'Class Roster Builder', desc: 'Full student rosters (name, email, section, last access, grade) for all your courses in a term.', course: false, fn: runClassRosterBuilder },
    { label: 'At-Risk Early Warning Dashboard', desc: 'Flags students inactive 7+ days, with a grade below 70%, or 2+ missing assignments \u2014 across a term.', course: false, fn: runAtRiskDashboard },
    { label: 'Peer Comparison Snapshot', desc: 'Anonymized comparison of one student vs. the rest of this course.', course: true, fn: runPeerComparison },
    { label: 'Re-Engagement Email Drafter', desc: 'Drafts check-in emails for students flagged inactive 7+ days, below 70%, or with 2+ missing assignments, across a term.', course: false, fn: runReengagementDrafter },
    { label: 'Weekly Announcement Auto-Drafter', desc: 'Drafts this week\u2019s announcement for this course.', course: true, fn: runWeeklyAnnouncementDrafter },
    { label: 'Discussion to Word Exporter', desc: 'Exports the current discussion thread \u2014 original post and all replies \u2014 as a .docx file.', course: false, discussion: true, fn: runDiscussionExport },
    { label: 'Canvas Grade Report', desc: 'Active + Inactive students across a term, with grades, activity, last submission date, live search, and Excel export.', course: false, fn: runGradeReport },
    { label: 'Pre-No-Show Report', desc: 'Flags students with zero course participation (no posts, submissions, quizzes, or LTI activity) across a term, for financial-aid no-show reporting. Draft outreach opens in Canvas Message for you to review and send.', course: false, fn: runPreNoShowReport }
  ];

  removeOverlay('canvas-toolkit-menu');
  var overlay = document.createElement('div');
  overlay.id = 'canvas-toolkit-menu';
  overlay.style.cssText = 'position:fixed;top:16px;left:16px;right:16px;bottom:16px;background:#fff;z-index:999999;overflow:auto;padding:24px;border:1px solid #ccc;border-radius:6px;box-shadow:0 4px 24px rgba(0,0,0,0.35);font-family:-apple-system,Helvetica,Arial,sans-serif;color:#222;';

  var html = '<div style="text-align:right;margin-bottom:14px;">' +
    '<button id="ctk-close-btn" style="padding:8px 14px;cursor:pointer;border:1px solid #999;border-radius:4px;background:#f5f5f5;">Close</button>' +
    '</div>';
  html += '<h2 style="margin-top:0;">Canvas Toolkit</h2>' +
    '<p style="color:#666;font-size:13px;margin-bottom:16px;">Pick a tool to run. Tools marked <strong>(this course)</strong> need to be run from inside the course you want to use them on \u2014 the rest work from anywhere and let you pick a term.</p>';

  tools.forEach(function(t, i){
    html += '<button class="ctk-tool-btn" data-idx="' + i + '" style="display:block;width:100%;text-align:left;padding:12px 14px;margin-bottom:10px;cursor:pointer;border:1px solid #999;border-radius:4px;background:#f5f5f5;">' +
      '<div style="font-weight:600;">' + esc(t.label) +
        (t.course ? ' <span style="font-weight:400;color:#a66;font-size:12px;">(this course)</span>' : '') +
        (t.discussion ? ' <span style="font-weight:400;color:#a66;font-size:12px;">(discussion page)</span>' : '') +
      '</div>' +
      '<div style="font-size:13px;color:#555;margin-top:2px;">' + esc(t.desc) + '</div>' +
      '</button>';
  });

  overlay.innerHTML = html;
  document.body.appendChild(overlay);

  document.getElementById('ctk-close-btn').onclick = function(){ overlay.remove(); };
  var buttons = overlay.querySelectorAll('.ctk-tool-btn');
  buttons.forEach(function(btn){
    btn.onclick = function(){
      var idx = parseInt(btn.getAttribute('data-idx'), 10);
      overlay.remove();
      tools[idx].fn();
    };
  });

})();