(function ($, sr) {

    // debouncing function from John Hann
    // http://unscriptable.com/index.php/2009/03/20/debouncing-javascript-methods/
    var debounce = function (func, threshold, execAsap) {
        var timeout;

        return function debounced() {
            var obj = this, args = arguments;

            function delayed() {
                if (!execAsap)
                    func.apply(obj, args);
                timeout = null;
            };

            if (timeout)
                clearTimeout(timeout);
            else if (execAsap)
                func.apply(obj, args);

            timeout = setTimeout(delayed, threshold || 100);
        };
    }
    // smartresize
    jQuery.fn[sr] = function (fn) {
        return fn ? this.bind('resize', debounce(fn)) : this.trigger(sr);
    };

})(jQuery, 'smartresize');


$.fn.excelSelect = function (options) {
    var allTables = this;
    allTables.each(function () {

        var api = {};
        var isSelecting = false;
        var showWireRect = true;
        var firstCell = {}, lastCell = {};
        var $table = $(this);
        $table.attr('tabindex', 0);

        var $wireRect = $('<div class="escel-wire-rect"><div class="escel-wire-extend"/></div>');

        if ($table.prop("tagName").toLowerCase() != 'table') {
            console.log("Can not apply escel table on " + this.prop("tagName") + "!!");
            return;
        }
        if ($table.length > 1) {
            console.log("Multiple tables not allowed!!");
            return;
        }


        var defaults = {
            selectedcellClass: "escel-selected-cell",
            firstCellSelectedClass: "escel-selected-cell first-escel-selected-cell",
            excelSelectTableClass: "escel-table",
            InjectCssIfIframe: true,
            IFrameFix_escelCssUrlPath: "",
            topCellClass: "escel-cell-t",
            rightCellClass: "escel-cell-r",
            leftCellClass: "escel-cell-l",
            bottomCellClass: "escel-cell-b",
            singleTableSelection: true,
            onSelectionEnd: function () {

            },
            onRightClick: function (rApi, cell, table, event) {
               
            }
        };
        defaults.allBorderClass = [defaults.topCellClass, defaults.rightCellClass, defaults.leftCellClass, defaults.bottomCellClass].join(' ');
        var settings = $.extend({}, defaults, options);
        if (this.ownerDocument && settings.InjectCssIfIframe && settings.IFrameFix_escelCssUrlPath.length > 0)//if iframe
        {
            var isInsert = true;
            var $head = $(this.ownerDocument).contents().find("head");
            if ($head.length == 0)
                $head = $('<head>');
            else {
                isInsert = $('link[href="' + settings.IFrameFix_escelCssUrlPath + '"]', $head).length <= 0;
            }

            if (isInsert)
                $head.append($("<link/>",
                    { rel: "stylesheet", href: settings.IFrameFix_escelCssUrlPath, type: "text/css" }));


        }

        $table.wrap($('<div class="escel-table-container"></div>'));
        var $tableContainer = $table.parent('.escel-table-container');
        $wireRect.appendTo($tableContainer);

        $table.addClass(settings.excelSelectTableClass);
        $tableContainer.disableSelection();


        var doc = this.ownerDocument || document;
        var win = doc.defaultView || doc.parentWindow;

        $(win).smartresize(function () {
            drawWireFrameOnSelectedCells();
        });

        var onMouseDown = function (event) {
            //$table.focus();
            if (event.which == 3)//if right click
            {
                if (isSelected(this)) {
                    settings.onRightClick.call(this, api, $table, event);
                    return;
                }
            }
            if (settings.singleTableSelection) {
                $(allTables).each(function (index, tbl) {
                    if (tbl != $table[0]) {
                        var api = $(tbl).data('escel-api');
                        api.clearSelection(true);
                    }
                });
            }
            isSelecting = true;
            if (event.shiftKey) {
                lastCell = $(this).cellPos();
            }
            else {
                if (!event.ctrlKey) {
                    clearSelection();
                    showWireRect = true;
                }

                else
                    showWireRect = false;
                firstCell = $(this).cellPos();
                lastCell = firstCell;
            }
            addRectPointsToSelection(firstCell, lastCell);

            if (event.which == 3)//if right click
            {
                settings.onRightClick.call(this, api, $table, event);
            }
        };
        var onMouseOver = function (event) {
            if (isSelecting) {
                if (!event.ctrlKey)
                    clearSelection();

                if (this == firstCell) {
                    addRectPointsToSelection(firstCell, firstCell);
                }
                else {
                    //var row = $(this).closest('tr');
                    var curCell = $(this).cellPos();
                    //curCell.rowIndex = $('tr', $table).index(row);
                    //curCell.cellIndex = $('td', row).index(this);
                    addRectPointsToSelection(firstCell, curCell);
                }
            }
        };
        var OnMouseUp = function (event) {
            endSelection();
        };
        var OnKeyUp = function (event) {

        }
        var OnKeyDown = function (event) {
            if ($table.is(":focus")) {
                var $nextCell = getFirstSelectedCell();
                var cellPt = $nextCell.cellPos();
                if (cellPt) {
                    cellPt.left++;
                    //cellPt = getNextCellPos(cellPt);
                    //if (cellPt) {
                        clearSelection();
                        addRectPointsToSelection(cellPt, cellPt);
                    //}
                }
                event.preventDefault();
            }
        }

        var endSelection = function () {
            if (isSelecting) {
                isSelecting = false;
                settings.onSelectionEnd.call(this, $table, $table.data('escel-api'));
            }
        };
        ///$table.on('keyup', OnKeyUp);
        //$table.on('keydown', OnKeyDown);
        $('td,th', $table)
            .on('mousedown', onMouseDown)
            .on('mouseover', onMouseOver)
            .on('mouseup', OnMouseUp);

        //.on('click', OnClick);

        $(document).on('mouseup', OnMouseUp);

        var getSelectedCells = function () {
            return $('td.' + settings.selectedcellClass + ',th.' + settings.selectedcellClass, $table);
        };
        var getFirstSelectedCell = function () {
            return $('td.' + settings.selectedcellClass + ',th.' + settings.selectedcellClass, $table).first();
        };
        var getLastSelectedCell = function () {
            return $('td.' + settings.selectedcellClass + ',th.' + settings.selectedcellClass, $table).last();
        };
        var destroy = function () {

            $(document).enableSelection();
            if ($table.get(0).ownerDocument)
                $($table.get(0).ownerDocument).enableSelection();

            $tableContainer.enableSelection();
            $table.removeClass(settings.excelSelectTableClass);
            $('td,th', $table).removeClass(settings.selectedcellClass);
            //$wireRect.remove();
            $(".escel-wire-rect", $tableContainer).remove();

            $table.unwrap();
            $table.removeData('escel-api');

            clearSelection();
            //$table.off('keyup', OnKeyUp);
            //$table.off('keydown', OnKeyDown);
            $('td,th', $table)
                .off('mousedown', onMouseDown)
                .off('mouseover', onMouseOver)
                .off('mouseup', OnMouseUp)
            //$('td,th', $table)
            //.off('mousedown', onMouseDown)
            //.off('mouseover', onMouseOver)
            //.off('mouseup', OnMouseUp);
            //.off('click', OnClick);
            $(document).off('mouseup', OnMouseUp);

        };

        var isSelected = function ($elm) {
            $elm = $($elm);
            return $elm.hasClass(settings.selectedcellClass);

        };
        var clearSelection = function (clearWireRect) {
            $('td', $table).removeClass(settings.firstCellSelectedClass + " " + settings.allBorderClass);
            if (clearWireRect)
                $wireRect.hide();

        };
        var addRectPointsToSelection = function (point1, poin2) {

            var maxRow = Math.max(point1.top, poin2.top);
            var minRow = Math.min(point1.top, poin2.top);
            var maxCol = Math.max(point1.left, poin2.left);
            var minCol = Math.min(point1.left, poin2.left);

            $('td,th', $table)
                .each(function (index, cell) {
                    var cellPos = $(cell).cellPos();
                    for (var iRowIndex = minRow; iRowIndex <= maxRow; iRowIndex++) {
                        for (var iCellIndex = minCol; iCellIndex <= maxCol; iCellIndex++) {

                            if (cellPos.top == iRowIndex && cellPos.left == iCellIndex) {
                                if (this.rowSpan > 1) {
                                    maxRow = Math.max(maxRow, cellPos.top + (this.rowSpan - 1));
                                }
                                if (this.colSpan > 1) {
                                    maxCol = Math.max(maxCol, cellPos.left + (this.colSpan - 1));
                                }
                            }
                        }
                    }
                })
                .each(function (index, cell) {

                    var $cell = $(cell);
                    var cellPos = $cell.cellPos();

                    for (var iRowIndex = minRow; iRowIndex <= maxRow; iRowIndex++) {
                        for (var iCellIndex = minCol; iCellIndex <= maxCol; iCellIndex++) {
                            if (cellPos.top == iRowIndex && cellPos.left == iCellIndex) {

                                var classToAdd = settings.selectedcellClass;
                                if (iRowIndex == minRow && iCellIndex == minCol) {
                                    classToAdd = settings.firstCellSelectedClass;
                                    // $leftTop = $cell;
                                }

                                //if (iRowIndex == maxRow && iCellIndex == maxCol)
                                //    $bottomRight = $cell;

                                $cell.addClass(classToAdd);
                            }
                        }
                    }
                });

            drawWireFrameOnSelectedCells();
        };
        var drawWireFrameOnSelectedCells = function () {
            var $selected = getSelectedCells();
            var $leftTop = null, $bottomRight = null;
            if ($selected.length > 0) {

                $leftTop = $($selected.get(0));
                $bottomRight = $($selected.get(-1));
            }


            if ($leftTop && $bottomRight) {
                var width = $bottomRight.offset().left - $leftTop.offset().left + $bottomRight.width();
                var height = $bottomRight.offset().top - $leftTop.offset().top + $bottomRight.height();
                drawWireFrame($leftTop.position().left, $leftTop.position().top, width, height);
            }

        };
        var drawWireFrame = function (left, top, width, height) {
            if (showWireRect) {
                $wireRect
                    .css(
                    {
                        left: left,
                        top: top,
                        width: width,
                        height: height
                    })
                    .show();
            }
            else {
                $wireRect.hide();
            }
        };
        var getNextCellPos = function (afterCell) {
            $('td,th', $table).each(function (index, cell) {
                var $cell = $(cell);
                var cellPos = $cell.cellPos();
                if (cellPos.top == afterCell.top && cellPos.left == afterCell.left)
                    return cellPos;
            });
            return null;
        };
        api = {
            getSelectedCells: getSelectedCells,
            destroy: destroy,
            clearSelection: clearSelection
        }
        $table.data('escel-api', api);
    });
    return this;
};

(function ($) {
   
    function scanTable($table) {
        var m = [];
        var lastCellPos = null;
        $table.children("tr").each(function (y, row) {
            $(row).children("td, th").each(function (x, cell) {
                var $cell = $(cell),
                    cspan = $cell.attr("colspan") | 0,
                    rspan = $cell.attr("rowspan") | 0,
                    tx, ty;
                cspan = cspan ? cspan : 1;
                rspan = rspan ? rspan : 1;
                for (; m[y] && m[y][x]; ++x); 
                for (tx = x; tx < x + cspan; ++tx) {  
                    for (ty = y; ty < y + rspan; ++ty) {
                        if (!m[ty]) {  
                            m[ty] = [];
                        }
                        m[ty][tx] = true;
                    }
                }
                var pos = { top: y, left: x };
                $cell.data("cellPos", pos);
               
                lastCellPos = pos;
            });
        });
        $table.data('lastCellPos', lastCellPos);

    };

   
    $.fn.cellPos = function (rescan) {
        var pos = null;
        var $cell = this.first(),
            pos = $cell.data("cellPos");
        if (!pos || rescan) {
            var $table = $cell.closest("table, thead, tbody, tfoot");
            scanTable($table);
        }
        pos = $cell.data("cellPos");
        return pos;
    };
    jQuery.fn.extend({
        unwrapInner: function (selector) {
            return this.each(function () {
                var t = this,
                    c = $(t).children(selector);
                if (c.length === 1) {
                    c.contents().appendTo(t);
                    c.remove();
                }
            });
        }
    });
})(jQuery);
(function ($) {
    $.fn.disableSelection = function () {
        return this.bind(($.support.selectstart ? "selectstart" : "mousedown") +
            ".disableSelection", function (event) {
            event.preventDefault();
        });
    };

    $.fn.enableSelection = function () {
        return this.unbind('.disableSelection');
    };
})(jQuery);
$.fn.getCellByCellPos = function (left, top) {
    return this.filter(
        function () {
            var val = $(this).data('cellPos');
            return val.left == left && val.top == top;
        }
    );
};


