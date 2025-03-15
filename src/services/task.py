"""Task"""
import threading
from queue import Queue
from typing import Callable

class TaskService:
    """任务管理服务"""
    
    def __init__(self, max_workers: int = 4):
        self.task_queue = Queue()
        self.workers = []
        self._create_workers(max_workers)
        
    def _create_workers(self, num_workers: int):
        """创建工作线程"""
        for _ in range(num_workers):
            worker = threading.Thread(target=self._worker_loop)
            worker.daemon = True
            worker.start()
            self.workers.append(worker)
            
    def _worker_loop(self):
        """工作线程主循环"""
        while True:
            task, args, kwargs = self.task_queue.get()
            try:
                task(*args, **kwargs)
            except Exception as e:
                print(f"Task failed: {e}")
            finally:
                self.task_queue.task_done()
                
    def submit(self, task: Callable, *args, **kwargs):
        """提交任务
        
        Args:
            task: 要执行的任务函数
            args: 位置参数
            kwargs: 关键字参数
        """
        self.task_queue.put((task, args, kwargs))
        
    def wait_completion(self):
        """等待所有任务完成"""
        self.task_queue.join()