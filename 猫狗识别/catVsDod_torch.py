import torch
import torchvision
from torch.autograd import Variable
from torchvision import transforms
from torchvision import datasets

# 准备工作，指定每批次训练样本，选择GPU进行计算
BATCH_SIZE = 64
DEVICE = torch.device('cuda:0' if torch.cuda.is_available() else 'cpu')
print(DEVICE)

# 图像预处理，转化为28*28像素
transform = transforms.Compose([
    transforms.Resize((28, 28)),
    transforms.ToTensor(),

])

transform_test = transforms.Compose([
    transforms.Resize((28, 28)),
    transforms.ToTensor(),

])

# 加载训练集和测试集
dataset_train = datasets.ImageFolder('E:/pyTest/dog vs cat/dataset/training_set', transform)
dataset_test = datasets.ImageFolder('E:/pyTest/dog vs cat/dataset/test_set', transform_test)
# print(dataset_train.imgs)

train_loader = torch.utils.data.DataLoader(dataset_train, batch_size=BATCH_SIZE, shuffle=True)
test_loader = torch.utils.data.DataLoader(dataset_test, batch_size=BATCH_SIZE, shuffle=False)

# 设置模型参数
class Net(torch.nn.Module):
    def __init__(self):
        super(Net, self).__init__()
        self.conv1 = torch.nn.Sequential(
            torch.nn.Conv2d(3, 32, 3, 1, 1),
            torch.nn.ReLU(),
            torch.nn.MaxPool2d(2))  # 卷积层
        self.conv2 = torch.nn.Sequential(
            torch.nn.Conv2d(32, 64, 3, 1, 1),
            torch.nn.ReLU(),
            torch.nn.MaxPool2d(2))
        self.conv3 = torch.nn.Sequential(
            torch.nn.Conv2d(64, 64, 3, 1, 1),
            torch.nn.ReLU(),
            torch.nn.MaxPool2d(2))
        self.dense = torch.nn.Sequential(
            torch.nn.Linear(64 * 3 * 3, 128),
            torch.nn.ReLU(),
            torch.nn.Linear(128, 10))   # 全连接层

    # 前向传播
    def forward(self, x):
        conv1_out = self.conv1(x)
        conv2_out = self.conv2(conv1_out)
        conv3_out = self.conv3(conv2_out)
        res = conv3_out.view(conv3_out.size(0), -1)
        out = self.dense(res)
        return out


model = Net().to(DEVICE)
print(model)

# 定义优化器
print(model.parameters())
optimizer = torch.optim.Adam(model.parameters())
loss_func = torch.nn.CrossEntropyLoss()

# 训练
for epoch in range(10):
    print(' epoch {}'.format(epoch + 1))
    # training--------------
    train_loss = 0.
    train_acc = 0.
    for batch_x, batch_y in train_loader:
        batch_x, batch_y = Variable(batch_x).to(DEVICE), Variable(batch_y).to(DEVICE)
        out = model(batch_x)
        loss = loss_func(out, batch_y)
        train_loss += loss.item()
        pred = torch.max(out, 1)[1]
        train_correct = (pred == batch_y).sum()
        train_acc += train_correct.item()
        optimizer.zero_grad()
        loss.backward()
        optimizer.step()
    print(len(dataset_train))
    print('Train Loss: {:.6f}，Acc: {:.6f} '.format(train_loss / (len(dataset_train)), train_acc / (len(dataset_train))))
    # evaluation-----------------
    model.eval()
    eval_loss = 0.
    eval_acc = 0.

torch.save(model, 'catVsDog.pth')

# 测试
for batch_x, batch_y in test_loader:
    with torch.no_grad():
        batch_x, batch_y = batch_x.to(DEVICE), batch_y.to(DEVICE)
        out = model(batch_x)
        loss = loss_func(out, batch_y)
        eval_loss += loss.item()
        pred = torch.max(out, 1)[1]
        num_correct = (pred == batch_y).sum()
        eval_acc += num_correct.item()

print('Test Loss: {:.6f}，Acc: {:.6f} '.format(eval_loss / (len(dataset_test)), eval_acc / (len(dataset_test))))


