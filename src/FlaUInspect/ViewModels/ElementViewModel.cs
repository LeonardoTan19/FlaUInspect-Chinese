using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using FlaUI.Core;
using FlaUI.Core.AutomationElements;
using FlaUI.Core.Conditions;
using FlaUI.Core.Definitions;
using FlaUI.Core.Identifiers;
using FlaUI.Core.Patterns;
using FlaUI.Core.Tools;
using FlaUI.UIA3.Identifiers;
using FlaUInspect.Core;

namespace FlaUInspect.ViewModels
{
    public class ElementViewModel : ObservableObject
    {
        public event Action<ElementViewModel> SelectionChanged;

        public ElementViewModel(AutomationElement automationElement)
        {
            AutomationElement = automationElement;
            Children = new ExtendedObservableCollection<ElementViewModel>();
            ItemDetails = new ExtendedObservableCollection<DetailGroupViewModel>();
        }

        public AutomationElement AutomationElement { get; }

        public bool IsSelected
        {
            get { return GetProperty<bool>(); }
            set
            {
                try
                {
                    if (value)
                    {
                        ElementHighlighter.HighlightElement(AutomationElement);

                        // Async load details
                        var unused = Task.Run(() =>
                        {
                            var details = LoadDetails();
                            return details;
                        }).ContinueWith(items =>
                        {
                            if (items.IsFaulted)
                            {
                                if (items.Exception != null)
                                {
                                    MessageBox.Show(items.Exception.ToString());
                                }
                            }
                            ItemDetails.Reset(items.Result);
                        }, TaskScheduler.FromCurrentSynchronizationContext());

                        // Fire the selection event
                        SelectionChanged?.Invoke(this);
                    }

                    SetProperty(value);
                }
                catch (Exception ex)
                {
                    Console.Write(ex.ToString());
                }
            }
        }

        public bool IsExpanded
        {
            get { return GetProperty<bool>(); }
            set
            {
                SetProperty(value);
                if (value)
                {
                    LoadChildren(true);
                }
            }
        }

        public string Name => NormalizeString(AutomationElement.Properties.Name.ValueOrDefault);

        public string AutomationId => NormalizeString(AutomationElement.Properties.AutomationId.ValueOrDefault);

        public ControlType ControlType => AutomationElement.Properties.ControlType.TryGetValue(out ControlType value) ? value : ControlType.Custom;

        public ExtendedObservableCollection<ElementViewModel> Children { get; set; }

        public ExtendedObservableCollection<DetailGroupViewModel> ItemDetails { get; set; }

        public string XPath => Debug.GetXPathToElement(AutomationElement);

        public void LoadChildren(bool loadInnerChildren)
        {
            foreach (var child in Children)
            {
                child.SelectionChanged -= SelectionChanged;
            }

            var childrenViewModels = new List<ElementViewModel>();
            try
            {
                foreach (var child in AutomationElement.FindAllChildren())
                {
                    var childViewModel = new ElementViewModel(child);
                    childViewModel.SelectionChanged += SelectionChanged;
                    childrenViewModels.Add(childViewModel);

                    if (loadInnerChildren)
                    {
                        childViewModel.LoadChildren(false);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"异常: {ex.Message}");
            }

            Children.Reset(childrenViewModels);
        }

        private List<DetailGroupViewModel> LoadDetails()
        {
            var detailGroups = new List<DetailGroupViewModel>();
            var cacheRequest = new CacheRequest();
            cacheRequest.TreeScope = TreeScope.Element;
            cacheRequest.Add(AutomationElement.Automation.PropertyLibrary.Element.AutomationId);
            cacheRequest.Add(AutomationElement.Automation.PropertyLibrary.Element.Name);
            cacheRequest.Add(AutomationElement.Automation.PropertyLibrary.Element.ClassName);
            cacheRequest.Add(AutomationElement.Automation.PropertyLibrary.Element.ControlType);
            cacheRequest.Add(AutomationElement.Automation.PropertyLibrary.Element.LocalizedControlType);
            cacheRequest.Add(AutomationElement.Automation.PropertyLibrary.Element.FrameworkId);
            cacheRequest.Add(AutomationElement.Automation.PropertyLibrary.Element.ProcessId);
            cacheRequest.Add(AutomationElement.Automation.PropertyLibrary.Element.IsEnabled);
            cacheRequest.Add(AutomationElement.Automation.PropertyLibrary.Element.IsOffscreen);
            cacheRequest.Add(AutomationElement.Automation.PropertyLibrary.Element.BoundingRectangle);
            cacheRequest.Add(AutomationElement.Automation.PropertyLibrary.Element.HelpText);
            cacheRequest.Add(AutomationElement.Automation.PropertyLibrary.Element.IsPassword);
            cacheRequest.Add(AutomationElement.Automation.PropertyLibrary.Element.NativeWindowHandle);
            using (cacheRequest.Activate())
            {
                var elementCached = AutomationElement.FindFirst(TreeScope.Element, TrueCondition.Default);
                if (elementCached != null)
                {
                    // Element identification
                    var identification = new List<IDetailViewModel>
                    {
                        DetailViewModel.FromAutomationProperty("自动化ID", elementCached.Properties.AutomationId),
                        DetailViewModel.FromAutomationProperty("名称", elementCached.Properties.Name),
                        DetailViewModel.FromAutomationProperty("类名", elementCached.Properties.ClassName),
                        DetailViewModel.FromAutomationProperty("控件类型", elementCached.Properties.ControlType),
                        DetailViewModel.FromAutomationProperty("本地化控件类型", elementCached.Properties.LocalizedControlType),
                        new DetailViewModel("框架类型", elementCached.FrameworkType.ToString()),
                        DetailViewModel.FromAutomationProperty("框架ID", elementCached.Properties.FrameworkId),
                        DetailViewModel.FromAutomationProperty("进程ID", elementCached.Properties.ProcessId),
                    };
                    detailGroups.Add(new DetailGroupViewModel("标识信息", identification));

                    // Element details
                    var details = new List<DetailViewModel>
                    {
                        DetailViewModel.FromAutomationProperty("是否启用", elementCached.Properties.IsEnabled),
                        DetailViewModel.FromAutomationProperty("是否屏幕外", elementCached.Properties.IsOffscreen),
                        DetailViewModel.FromAutomationProperty("边界矩形", elementCached.Properties.BoundingRectangle),
                        DetailViewModel.FromAutomationProperty("帮助文本", elementCached.Properties.HelpText),
                        DetailViewModel.FromAutomationProperty("是否密码", elementCached.Properties.IsPassword)
                    };
                    // Special handling for NativeWindowHandle
                    var nativeWindowHandle = elementCached.Properties.NativeWindowHandle.ValueOrDefault;
                    var nativeWindowHandleString = "不支持";
                    if (nativeWindowHandle != default(IntPtr))
                    {
                        nativeWindowHandleString = String.Format("{0} ({0:X8})", nativeWindowHandle.ToInt32());
                    }
                    details.Add(new DetailViewModel("本机窗口句柄", nativeWindowHandleString));
                    detailGroups.Add(new DetailGroupViewModel("详细信息", details));
                }
            }

            // Pattern details
            var allSupportedPatterns = AutomationElement.GetSupportedPatterns();
            var allPatterns = AutomationElement.Automation.PatternLibrary.AllForCurrentFramework;
            var patterns = new List<DetailViewModel>();
            foreach (var pattern in allPatterns)
            {
                var hasPattern = allSupportedPatterns.Contains(pattern);
                patterns.Add(new DetailViewModel(pattern.Name + "模式", hasPattern ? "是" : "否") { Important = hasPattern });
            }
            detailGroups.Add(new DetailGroupViewModel("模式支持", patterns));

            // GridItemPattern
            if (allSupportedPatterns.Contains(AutomationElement.Automation.PatternLibrary.GridItemPattern))
            {
                var pattern = AutomationElement.Patterns.GridItem.Pattern;
                var patternDetails = new List<DetailViewModel>
                {
                    DetailViewModel.FromAutomationProperty("列", pattern.Column),
                    DetailViewModel.FromAutomationProperty("列跨度", pattern.ColumnSpan),
                    DetailViewModel.FromAutomationProperty("行", pattern.Row),
                    DetailViewModel.FromAutomationProperty("行跨度", pattern.RowSpan),
                    DetailViewModel.FromAutomationProperty("包含网格", pattern.ContainingGrid)
                };
                detailGroups.Add(new DetailGroupViewModel("网格项模式", patternDetails));
            }
            // GridPattern
            if (allSupportedPatterns.Contains(AutomationElement.Automation.PatternLibrary.GridPattern))
            {
                var pattern = AutomationElement.Patterns.Grid.Pattern;
                var patternDetails = new List<DetailViewModel>
                {
                    DetailViewModel.FromAutomationProperty("列数", pattern.ColumnCount),
                    DetailViewModel.FromAutomationProperty("行数", pattern.RowCount)
                };
                detailGroups.Add(new DetailGroupViewModel("网格模式", patternDetails));
            }
            // LegacyIAccessiblePattern
            if (allSupportedPatterns.Contains(AutomationElement.Automation.PatternLibrary.LegacyIAccessiblePattern))
            {
                var pattern = AutomationElement.Patterns.LegacyIAccessible.Pattern;
                var patternDetails = new List<DetailViewModel>
                {
                   DetailViewModel.FromAutomationProperty("名称", pattern.Name),
                   new DetailViewModel("状态", AccessibilityTextResolver.GetStateText(pattern.State.ValueOrDefault)),
                   new DetailViewModel("角色", AccessibilityTextResolver.GetRoleText(pattern.Role.ValueOrDefault)),
                   DetailViewModel.FromAutomationProperty("值", pattern.Value),
                   DetailViewModel.FromAutomationProperty("子项ID", pattern.ChildId),
                   DetailViewModel.FromAutomationProperty("默认操作", pattern.DefaultAction),
                   DetailViewModel.FromAutomationProperty("描述", pattern.Description),
                   DetailViewModel.FromAutomationProperty("帮助", pattern.Help),
                   DetailViewModel.FromAutomationProperty("键盘快捷键", pattern.KeyboardShortcut),
                   DetailViewModel.FromAutomationProperty("选择", pattern.Selection)
                };
                detailGroups.Add(new DetailGroupViewModel("传统IAccessible模式", patternDetails));
            }
            // RangeValuePattern
            if (allSupportedPatterns.Contains(AutomationElement.Automation.PatternLibrary.RangeValuePattern))
            {
                var pattern = AutomationElement.Patterns.RangeValue.Pattern;
                var patternDetails = new List<DetailViewModel>
                {
                   DetailViewModel.FromAutomationProperty("是否只读", pattern.IsReadOnly),
                   DetailViewModel.FromAutomationProperty("小步长", pattern.SmallChange),
                   DetailViewModel.FromAutomationProperty("大步长", pattern.LargeChange),
                   DetailViewModel.FromAutomationProperty("最小值", pattern.Minimum),
                   DetailViewModel.FromAutomationProperty("最大值", pattern.Maximum),
                   DetailViewModel.FromAutomationProperty("当前值", pattern.Value)
                };
                detailGroups.Add(new DetailGroupViewModel("范围值模式", patternDetails));
            }
            // ScrollPattern
            if (allSupportedPatterns.Contains(AutomationElement.Automation.PatternLibrary.ScrollPattern))
            {
                var pattern = AutomationElement.Patterns.Scroll.Pattern;
                var patternDetails = new List<DetailViewModel>
                {
                    DetailViewModel.FromAutomationProperty("水平滚动百分比", pattern.HorizontalScrollPercent),
                    DetailViewModel.FromAutomationProperty("水平视图大小", pattern.HorizontalViewSize),
                    DetailViewModel.FromAutomationProperty("可水平滚动", pattern.HorizontallyScrollable),
                    DetailViewModel.FromAutomationProperty("垂直滚动百分比", pattern.VerticalScrollPercent),
                    DetailViewModel.FromAutomationProperty("垂直视图大小", pattern.VerticalViewSize),
                    DetailViewModel.FromAutomationProperty("可垂直滚动", pattern.VerticallyScrollable)
                };
                detailGroups.Add(new DetailGroupViewModel("滚动模式", patternDetails));
            }
            // SelectionItemPattern
            if (allSupportedPatterns.Contains(AutomationElement.Automation.PatternLibrary.SelectionItemPattern))
            {
                var pattern = AutomationElement.Patterns.SelectionItem.Pattern;
                var patternDetails = new List<DetailViewModel>
                {
                    DetailViewModel.FromAutomationProperty("是否选中", pattern.IsSelected),
                    DetailViewModel.FromAutomationProperty("选择容器", pattern.SelectionContainer)
                };
                detailGroups.Add(new DetailGroupViewModel("选择项模式", patternDetails));
            }
            // SelectionPattern
            if (allSupportedPatterns.Contains(AutomationElement.Automation.PatternLibrary.SelectionPattern))
            {
                var pattern = AutomationElement.Patterns.Selection.Pattern;
                var patternDetails = new List<DetailViewModel>
                {
                    DetailViewModel.FromAutomationProperty("选择", pattern.Selection),
                    DetailViewModel.FromAutomationProperty("可多选", pattern.CanSelectMultiple),
                    DetailViewModel.FromAutomationProperty("需要选择", pattern.IsSelectionRequired)
                };
                detailGroups.Add(new DetailGroupViewModel("选择模式", patternDetails));
            }
            // TableItemPattern
            if (allSupportedPatterns.Contains(AutomationElement.Automation.PatternLibrary.TableItemPattern))
            {
                var pattern = AutomationElement.Patterns.TableItem.Pattern;
                var patternDetails = new List<DetailViewModel>
                {
                    DetailViewModel.FromAutomationProperty("列标题项", pattern.ColumnHeaderItems),
                    DetailViewModel.FromAutomationProperty("行标题项", pattern.RowHeaderItems)
                };
                detailGroups.Add(new DetailGroupViewModel("表格项模式", patternDetails));
            }
            // TablePattern
            if (allSupportedPatterns.Contains(AutomationElement.Automation.PatternLibrary.TablePattern))
            {
                var pattern = AutomationElement.Patterns.Table.Pattern;
                var patternDetails = new List<DetailViewModel>
                {
                    DetailViewModel.FromAutomationProperty("列标题项", pattern.ColumnHeaders),
                    DetailViewModel.FromAutomationProperty("行标题项", pattern.RowHeaders),
                    DetailViewModel.FromAutomationProperty("行或列主要", pattern.RowOrColumnMajor)
                };
                detailGroups.Add(new DetailGroupViewModel("表格模式", patternDetails));
            }
            // TextPattern
            if (allSupportedPatterns.Contains(AutomationElement.Automation.PatternLibrary.TextPattern))
            {
                var pattern = AutomationElement.Patterns.Text.Pattern;

                // TODO: This can in the future be replaced with automation.MixedAttributeValue
                object mixedValue = AutomationElement.AutomationType == AutomationType.UIA2
                    ? System.Windows.Automation.TextPattern.MixedAttributeValue
                    : ((FlaUI.UIA3.UIA3Automation)AutomationElement.Automation).NativeAutomation.ReservedMixedAttributeValue;

                var foreColor = GetTextAttribute<int>(pattern, TextAttributes.ForegroundColor, mixedValue, (x) =>
                {
                    return $"{System.Drawing.Color.FromArgb(x)} ({x})";
                });
                var backColor = GetTextAttribute<int>(pattern, TextAttributes.BackgroundColor, mixedValue, (x) =>
                {
                    return $"{System.Drawing.Color.FromArgb(x)} ({x})";
                });
                var fontName = GetTextAttribute<string>(pattern, TextAttributes.FontName, mixedValue, (x) =>
                {
                    return $"{x}";
                });
                var fontSize = GetTextAttribute<double>(pattern, TextAttributes.FontSize, mixedValue, (x) =>
                {
                    return $"{x}";
                });
                var fontWeight = GetTextAttribute<int>(pattern, TextAttributes.FontWeight, mixedValue, (x) =>
                {
                    return $"{x}";
                });

                var patternDetails = new List<DetailViewModel>
                {
                    new DetailViewModel("前景色", foreColor),
                    new DetailViewModel("背景色", backColor),
                    new DetailViewModel("字体名称", fontName),
                    new DetailViewModel("字体大小", fontSize),
                    new DetailViewModel("字体粗细", fontWeight),
                };
                detailGroups.Add(new DetailGroupViewModel("文本模式", patternDetails));
            }
            // TogglePattern
            if (allSupportedPatterns.Contains(AutomationElement.Automation.PatternLibrary.TogglePattern))
            {
                var pattern = AutomationElement.Patterns.Toggle.Pattern;
                var patternDetails = new List<DetailViewModel>
                {
                    DetailViewModel.FromAutomationProperty("切换状态", pattern.ToggleState)
                };
                detailGroups.Add(new DetailGroupViewModel("切换模式", patternDetails));
            }
            // ValuePattern
            if (allSupportedPatterns.Contains(AutomationElement.Automation.PatternLibrary.ValuePattern))
            {
                var pattern = AutomationElement.Patterns.Value.Pattern;
                var patternDetails = new List<DetailViewModel>
                {
                    DetailViewModel.FromAutomationProperty("是否只读", pattern.IsReadOnly),
                    DetailViewModel.FromAutomationProperty("值", pattern.Value)
                };
                detailGroups.Add(new DetailGroupViewModel("值模式", patternDetails));
            }
            // WindowPattern
            if (allSupportedPatterns.Contains(AutomationElement.Automation.PatternLibrary.WindowPattern))
            {
                var pattern = AutomationElement.Patterns.Window.Pattern;
                var patternDetails = new List<DetailViewModel>
                {
                    DetailViewModel.FromAutomationProperty("是否模态", pattern.IsModal),
                    DetailViewModel.FromAutomationProperty("是否置顶", pattern.IsTopmost),
                    DetailViewModel.FromAutomationProperty("可最小化", pattern.CanMinimize),
                    DetailViewModel.FromAutomationProperty("可最大化", pattern.CanMaximize),
                    DetailViewModel.FromAutomationProperty("窗口视觉状态", pattern.WindowVisualState),
                    DetailViewModel.FromAutomationProperty("窗口交互状态", pattern.WindowInteractionState)
                };
                detailGroups.Add(new DetailGroupViewModel("窗口模式", patternDetails));
            }

            return detailGroups;
        }

        private string GetTextAttribute<T>(ITextPattern pattern, TextAttributeId textAttribute, object mixedValue, Func<T, string> func)
        {
            var value = pattern.DocumentRange.GetAttributeValue(textAttribute);

            if (value == mixedValue)
            {
                return "混合";
            }
            else if (value == AutomationElement.Automation.NotSupportedValue)
            {
                return "不支持";
            }
            else
            {
                try
                {
                    var converted = (T)value;
                    return func(converted);
                }
                catch
                {
                    return $"转换为 ${typeof(T)} 失败";
                }
            }
        }

        private string NormalizeString(string value)
        {
            if (String.IsNullOrEmpty(value))
            {
                return value;
            }
            return value.Replace(Environment.NewLine, " ").Replace('\r', ' ').Replace('\n', ' ');
        }
    }
}
