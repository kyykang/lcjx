/**
 * 营销平台流程绩效分析平台 - 图表功能实现
 * 这个文件包含所有图表的创建和数据处理逻辑
 */

// 全局变量存储数据和图表实例
let chartData = null;
let charts = {};

/**
 * 页面完全加载后初始化
 */
window.addEventListener('load', function() {
    console.log('页面完全加载完成，开始初始化...');
    console.log('Chart.js是否已加载:', typeof Chart !== 'undefined');
    
    // 检查canvas元素是否存在
    console.log('检查canvas元素是否存在:');
    console.log('flowRankingChart:', document.getElementById('flowRankingChart'));
    console.log('durationRankingChart:', document.getElementById('durationRankingChart'));
    console.log('salesDurationChart:', document.getElementById('salesDurationChart'));
    console.log('purchaseDurationChart:', document.getElementById('purchaseDurationChart'));
    console.log('projectDurationChart:', document.getElementById('projectDurationChart'));
    
    if (typeof Chart === 'undefined') {
        console.error('Chart.js库未加载');
        showErrorState('Chart.js库加载失败');
        return;
    }
    
    showLoadingState();
    loadData();
});

/**
 * 加载JSON数据文件
 */
async function loadData() {
    try {
        console.log('正在加载数据...');
        
        // 显示加载状态
        showLoadingState();
        
        // 获取数据文件
        const response = await fetch('chart_data.json');
        
        if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status}`);
        }
        
        chartData = await response.json();
        
        if (!chartData.success) {
            throw new Error(chartData.error || '数据加载失败');
        }
        
        console.log('数据加载成功:', chartData);
        
        // 初始化所有图表和统计信息
        initializeCharts();
        updateStatistics();
        
    } catch (error) {
        console.error('数据加载失败:', error);
        showErrorState(error.message);
    }
}

/**
 * 显示加载状态
 */
function showLoadingState() {
    const modules = document.querySelectorAll('.module .chart-container');
    modules.forEach(container => {
        container.innerHTML = '<div class="loading">正在加载数据</div>';
    });
}

/**
 * 显示错误状态
 */
function showErrorState(message) {
    const modules = document.querySelectorAll('.module .chart-container');
    modules.forEach(container => {
        container.innerHTML = `<div class="error">数据加载失败: ${message}</div>`;
    });
}

/**
 * 初始化所有图表
 */
function initializeCharts() {
    console.log('开始初始化图表...');
    console.log('chartData:', chartData);
    
    try {
        // 1. 发起流程数及完成数排名
        console.log('创建发起流程数排名图表...');
        createFlowRankingChart();
        
        // 2. 流程运行时长排名
        console.log('创建流程运行时长排名图表...');
        createDurationRankingChart();
        
        // 3. 销售类流程平均运行时长排名
        console.log('创建销售类流程图表...');
        createCategoryDurationChart('销售类流程', 'salesDurationChart');
        
        // 4. 采购类流程平均运行时长排名
        console.log('创建采购类流程图表...');
        createCategoryDurationChart('采购类流程', 'purchaseDurationChart');
        
        // 5. 项目&产品管理类流程平均运行时长排名
        console.log('创建项目管理类流程图表...');
        createCategoryDurationChart('项目&产品管理类流程', 'projectDurationChart');
        
        console.log('所有图表初始化完成');
    } catch (error) {
        console.error('图表初始化失败:', error);
        showErrorState(error.message);
    }
}

/**
 * 创建发起流程数及完成数排名图表
 */
function createFlowRankingChart() {
    const canvas = document.getElementById('flowRankingChart');
    if (!canvas) {
        console.error('找不到flowRankingChart canvas元素');
        return;
    }
    
    const ctx = canvas.getContext('2d');
    const data = chartData.data.flow_ranking;
    
    // 准备图表数据
    const labels = data.map(item => truncateText(item.模板名称, 10));
    const initiatedData = data.map(item => item.发起流程数 || 0);
    const completedData = data.map(item => item.完成流程数 || 0);
    
    charts.flowRanking = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: labels,
            datasets: [
                {
                    label: '发起流程数',
                    data: initiatedData,
                    backgroundColor: 'rgba(52, 152, 219, 0.8)',
                    borderColor: 'rgba(52, 152, 219, 1)',
                    borderWidth: 1
                },
                {
                    label: '完成流程数',
                    data: completedData,
                    backgroundColor: 'rgba(46, 204, 113, 0.8)',
                    borderColor: 'rgba(46, 204, 113, 1)',
                    borderWidth: 1
                }
            ]
        },
        options: {
            indexAxis: 'y',
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                title: {
                    display: true,
                    text: '发起流程数 vs 完成流程数 (前10名)',
                    font: { size: 16 }
                },
                legend: {
                    position: 'top'
                },
                tooltip: {
                    callbacks: {
                        title: function(context) {
                            const index = context[0].dataIndex;
                            return data[index].模板名称;
                        },
                        afterBody: function(context) {
                            const index = context[0].dataIndex;
                            const item = data[index];
                            return [
                                `完成率: ${((item.完成流程数 / item.发起流程数) * 100).toFixed(1)}%`,
                                `未完成: ${item.未结束流程数 || 0}`
                            ];
                        }
                    }
                }
            },
            scales: {
                y: {
                    beginAtZero: true,
                    title: {
                        display: true,
                        text: '流程数量'
                    }
                },
                x: {
                    title: {
                        display: true,
                        text: '流程名称'
                    }
                }
            }
        }
    });
}

/**
 * 创建流程运行时长排名图表
 */
function createDurationRankingChart() {
    const canvas = document.getElementById('durationRankingChart');
    if (!canvas) {
        console.error('找不到durationRankingChart canvas元素');
        return;
    }
    
    const ctx = canvas.getContext('2d');
    const data = chartData.data.duration_ranking;
    
    // 准备图表数据
    const labels = data.map(item => truncateText(item.模板名称, 10));
    const durations = data.map(item => item.平均运行时长_数值 || 0);
    
    // 根据时长生成渐变色
    const colors = durations.map(duration => {
        const intensity = Math.min(duration / Math.max(...durations), 1);
        return `rgba(231, 76, 60, ${0.3 + intensity * 0.7})`;
    });
    
    charts.durationRanking = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: labels,
            datasets: [{
                label: '平均运行时长(小时)',
                data: durations,
                backgroundColor: colors,
                borderColor: 'rgba(231, 76, 60, 1)',
                borderWidth: 1
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                title: {
                    display: true,
                    text: '流程平均运行时长排名 (前10名)',
                    font: { size: 16 }
                },
                legend: {
                    display: false
                },
                tooltip: {
                    callbacks: {
                        title: function(context) {
                            const index = context[0].dataIndex;
                            return data[index].模板名称;
                        },
                        label: function(context) {
                            const hours = context.parsed.x;
                            const days = Math.floor(hours / 24);
                            const remainingHours = hours % 24;
                            
                            if (days > 0) {
                                return `平均时长: ${days}天${remainingHours.toFixed(1)}小时`;
                            } else {
                                return `平均时长: ${hours.toFixed(1)}小时`;
                            }
                        },
                        afterBody: function(context) {
                            const index = context[0].dataIndex;
                            const item = data[index];
                            return [
                                `发起数: ${item.发起流程数 || 0}`,
                                `完成数: ${item.完成流程数 || 0}`
                            ];
                        }
                    }
                }
            },
            scales: {
                x: {
                    beginAtZero: true,
                    title: {
                        display: true,
                        text: '时长(小时)'
                    }
                },
                y: {
                    title: {
                        display: true,
                        text: '流程名称'
                    }
                }
            }
        }
    });
}

/**
 * 创建分类流程运行时长图表
 */
function createCategoryDurationChart(categoryName, canvasId) {
    const canvas = document.getElementById(canvasId);
    if (!canvas) {
        console.error(`找不到${canvasId} canvas元素`);
        return;
    }
    
    const ctx = canvas.getContext('2d');
    const data = chartData.data.category_rankings[categoryName] || [];
    
    if (data.length === 0) {
        document.getElementById(canvasId).parentElement.innerHTML = 
            '<div class="error">该分类暂无数据</div>';
        return;
    }
    
    // 准备图表数据
    const labels = data.map(item => truncateText(item.模板名称, 8));
    const durations = data.map(item => item.平均运行时长_数值 || 0);
    
    // 为不同分类使用不同的颜色主题
    const colorThemes = {
        '销售类流程': 'rgba(155, 89, 182, 0.8)',
        '采购类流程': 'rgba(241, 196, 15, 0.8)',
        '项目&产品管理类流程': 'rgba(26, 188, 156, 0.8)'
    };
    
    const backgroundColor = colorThemes[categoryName] || 'rgba(52, 152, 219, 0.8)';
    
    charts[canvasId] = new Chart(ctx, {
        type: 'doughnut',
        data: {
            labels: labels,
            datasets: [{
                data: durations,
                backgroundColor: labels.map((_, index) => {
                    const hue = (index * 360 / labels.length) % 360;
                    return `hsla(${hue}, 70%, 60%, 0.8)`;
                }),
                borderColor: '#fff',
                borderWidth: 2
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                title: {
                    display: true,
                    text: `${categoryName}时长分布`,
                    font: { size: 16 }
                },
                legend: {
                    position: 'right',
                    labels: {
                        boxWidth: 12,
                        font: { size: 10 }
                    }
                },
                tooltip: {
                    callbacks: {
                        title: function(context) {
                            const index = context[0].dataIndex;
                            return data[index].模板名称;
                        },
                        label: function(context) {
                            const hours = context.parsed;
                            const days = Math.floor(hours / 24);
                            const remainingHours = hours % 24;
                            
                            if (days > 0) {
                                return `时长: ${days}天${remainingHours.toFixed(1)}小时`;
                            } else {
                                return `时长: ${hours.toFixed(1)}小时`;
                            }
                        },
                        afterBody: function(context) {
                            const index = context[0].dataIndex;
                            const item = data[index];
                            const total = durations.reduce((sum, val) => sum + val, 0);
                            const percentage = ((item.平均运行时长_数值 / total) * 100).toFixed(1);
                            
                            return [
                                `占比: ${percentage}%`,
                                `发起数: ${item.发起流程数 || 0}`
                            ];
                        }
                    }
                }
            }
        }
    });
}

/**
 * 更新统计信息卡片
 */
function updateStatistics() {
    const rawData = chartData.data.raw_data;
    
    // 计算统计数据
    const totalFlows = rawData.length;
    const totalInitiated = rawData.reduce((sum, item) => sum + (item.发起流程数 || 0), 0);
    const totalCompleted = rawData.reduce((sum, item) => sum + (item.完成流程数 || 0), 0);
    
    // 计算平均时长（只计算有时长数据的流程）
    const validDurations = rawData.filter(item => item.平均运行时长_数值 > 0);
    const avgDuration = validDurations.length > 0 
        ? validDurations.reduce((sum, item) => sum + item.平均运行时长_数值, 0) / validDurations.length
        : 0;
    
    // 更新页面显示
    document.getElementById('totalFlows').textContent = totalFlows.toLocaleString();
    document.getElementById('totalInitiated').textContent = totalInitiated.toLocaleString();
    document.getElementById('totalCompleted').textContent = totalCompleted.toLocaleString();
    document.getElementById('avgDuration').textContent = avgDuration.toFixed(1);
    
    console.log('统计信息更新完成:', {
        totalFlows,
        totalInitiated,
        totalCompleted,
        avgDuration: avgDuration.toFixed(1)
    });
}

/**
 * 截断文本，避免标签过长
 */
function truncateText(text, maxLength) {
    if (!text) return '';
    return text.length > maxLength ? text.substring(0, maxLength) + '...' : text;
}

/**
 * 格式化时长显示
 */
function formatDuration(hours) {
    if (hours < 24) {
        return `${hours.toFixed(1)}小时`;
    } else {
        const days = Math.floor(hours / 24);
        const remainingHours = hours % 24;
        return `${days}天${remainingHours.toFixed(1)}小时`;
    }
}

/**
 * 窗口大小改变时重新调整图表
 */
window.addEventListener('resize', function() {
    Object.values(charts).forEach(chart => {
        if (chart && typeof chart.resize === 'function') {
            chart.resize();
        }
    });
});

/**
 * 导出图表为图片（可选功能）
 */
function exportChart(chartId, filename) {
    const chart = charts[chartId];
    if (chart) {
        const url = chart.toBase64Image();
        const link = document.createElement('a');
        link.download = filename || 'chart.png';
        link.href = url;
        link.click();
    }
}

// 在控制台提供导出功能
console.log('图表导出功能已加载。使用 exportChart("chartId", "filename.png") 来导出图表。');
console.log('可用的图表ID:', [
    'flowRanking',
    'durationRanking', 
    'salesDurationChart',
    'purchaseDurationChart',
    'projectDurationChart'
]);